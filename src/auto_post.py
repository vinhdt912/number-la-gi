"""
Đăng bài Threads (account Bac Day) kèm ảnh/video.

Excel `bac_day_vocab_data.xlsx` — Sheet1, cột:
  - content      : nội dung bài chính
  - sub_content  : nội dung reply (luồng trả lời bài chính)
  - image_url    : URL công khai của ảnh (JPEG/PNG)
  - video_url    : URL công khai của video
  - upload_time  : tự ghi vào ô sau khi đăng thành công (để bỏ qua bài đã đăng)

Sheet `info`:
  - time_break   : số phút chờ giữa 2 bài (thay CONST_TIME_SLEEP)

Media: cả ảnh + video → CAROUSEL; chỉ ảnh → IMAGE; chỉ video → VIDEO; không media → TEXT.
"""

import os
import time
from datetime import datetime, timedelta

import pandas as pd
import random
import requests
from dotenv import load_dotenv
from openpyxl import load_workbook

CONST_MAX_POST = 15 # số bài đăng tối đa trong 1 phiên
DEFAULT_TIME_SLEEP = 60  # phút — fallback nếu sheet info thiếu time_break
SESSION_START_HOUR = 9 # giờ đăng bài
SESSION_END_HOUR = 0 # giờ kết thúc phiên đăng bài
POST_BUFFER_SECONDS = 180 # thời gian chờ giữa 2 bài 

REQUIRED_COLUMNS = ("content", "sub_content", "image_url", "video_url") # các cột bắt buộc trong file Excel
THREADS_API_BASE = "https://graph.threads.net/v1.0" # base URL của API Threads
CONTAINER_WAIT_TIMEOUT = 300 # thời gian chờ container xử lý xong (cần cho VIDEO / CAROUSEL)
CONTAINER_POLL_INTERVAL = 15 # khoảng thời gian poll container xử lý xong (cần cho VIDEO / CAROUSEL)


def _is_blank(value):
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    return str(value).strip() in ("", "nan", "None")


def _clean_str(value):
    if _is_blank(value):
        return None
    return str(value).strip()


def get_session_bounds(now=None):
    """Phiên đăng: 09:00 hôm nay → 00:00 hôm sau."""
    now = now or datetime.now()
    session_start = now.replace(
        hour=SESSION_START_HOUR, minute=0, second=0, microsecond=0
    )
    session_end = now.replace(
        hour=SESSION_END_HOUR, minute=0, second=0, microsecond=0
    )

    if now.hour >= SESSION_START_HOUR:
        session_end += timedelta(days=1)
        return session_start, session_end
    if now.hour < SESSION_END_HOUR:
        session_start -= timedelta(days=1)
        return session_start, session_end
    return None, None


def wait_until_datetime(target: datetime):
    now = datetime.now()
    if now >= target:
        return
    wait_seconds = (target - now).total_seconds()
    print(f"🕒 Bây giờ: {now.strftime('%Y-%m-%d %H:%M:%S')}")
    print(
        f"⏳ Chờ đến {target.strftime('%Y-%m-%d %H:%M:%S')} "
        f"({wait_seconds / 60:.0f} phút)..."
    )
    time.sleep(wait_seconds)


def wait_until_session_start():
    while True:
        session_start, session_end = get_session_bounds()
        if session_start is not None:
            print(
                f"📅 Phiên đăng: {session_start.strftime('%d/%m %H:%M')} "
                f"→ {session_end.strftime('%d/%m %H:%M')}"
            )
            return session_start, session_end

        now = datetime.now()
        target = now.replace(
            hour=SESSION_START_HOUR, minute=0, second=0, microsecond=0
        )
        print("⏸️  Ngoài khung giờ đăng (00:00–09:00). Chờ đến 9:00...")
        wait_until_datetime(target)


def filter_session_posts(df, session_start, session_end):
    return df[
        (df["upload_time"] >= session_start) & (df["upload_time"] < session_end)
    ].copy()


def load_vocab_df(file_path):
    df = pd.read_excel(file_path, sheet_name="Sheet1")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Excel thiếu cột: {missing}")
    if "upload_time" not in df.columns:
        df["upload_time"] = pd.NaT
    df["upload_time"] = pd.to_datetime(df["upload_time"], errors="coerce")
    return df


def load_time_break(file_path):
    """Đọc time_break (phút) từ sheet info. Fallback DEFAULT_TIME_SLEEP nếu thiếu."""
    try:
        info_df = pd.read_excel(file_path, sheet_name="info")
    except Exception as exc:
        print(
            f"Không đọc được sheet info, dùng mặc định "
            f"{DEFAULT_TIME_SLEEP} phút: {exc}"
        )
        return DEFAULT_TIME_SLEEP

    if "time_break" not in info_df.columns or info_df.empty:
        print(
            f"Sheet info thiếu time_break, dùng mặc định "
            f"{DEFAULT_TIME_SLEEP} phút."
        )
        return DEFAULT_TIME_SLEEP

    value = info_df["time_break"].iloc[0]
    try:
        minutes = int(float(value))
        if minutes <= 0:
            raise ValueError("time_break phải > 0")
        return minutes
    except (TypeError, ValueError):
        print(
            f"time_break không hợp lệ ({value}), dùng mặc định "
            f"{DEFAULT_TIME_SLEEP} phút."
        )
        return DEFAULT_TIME_SLEEP


def write_upload_time(file_path, row_idx, uploaded_at):
    """Ghi upload_time vào đúng ô Excel, không ghi đè cả sheet."""
    workbook = load_workbook(file_path)
    worksheet = workbook["Sheet1"]
    header_cells = next(worksheet.iter_rows(min_row=1, max_row=1))
    col_idx = None
    for cell in header_cells:
        if cell.value == "upload_time":
            col_idx = cell.column
            break
    if col_idx is None:
        col_idx = (worksheet.max_column or 0) + 1
        worksheet.cell(row=1, column=col_idx, value="upload_time")

    excel_row = int(row_idx) + 2
    worksheet.cell(row=excel_row, column=col_idx, value=uploaded_at)
    workbook.save(file_path)


def seconds_until_session_end(session_end):
    return (session_end - datetime.now()).total_seconds()


def can_post_more(session_end, posted_count):
    if posted_count >= CONST_MAX_POST:
        return False
    if seconds_until_session_end(session_end) <= POST_BUFFER_SECONDS:
        return False
    return datetime.now() < session_end


def sleep_between_posts(session_end, file_path):
    remaining = seconds_until_session_end(session_end) - POST_BUFFER_SECONDS
    if remaining <= 0:
        return
    time_break = load_time_break(file_path)
    random_time_break = random.randint(time_break - 10, time_break + 10)
    time_sleep = min(random_time_break * 60, remaining)
    print(f"Đang đợi {time_sleep / 60:.0f} phút trước bài tiếp theo...")
    time.sleep(time_sleep)


def resolve_media(image_url=None, video_url=None):
    """Trả về (media_type, image_url, video_url). Cả hai URL → CAROUSEL."""
    image_url = _clean_str(image_url)
    video_url = _clean_str(video_url)
    if image_url and video_url:
        return "CAROUSEL", image_url, video_url
    if image_url:
        return "IMAGE", image_url, None
    if video_url:
        return "VIDEO", None, video_url
    return "TEXT", None, None


def create_threads_container(user_id, payload):
    url = f"{THREADS_API_BASE}/{user_id}/threads"
    response = requests.post(url, data=payload)
    data = response.json()
    if "id" not in data:
        print(f"Lỗi tạo container: {data}")
        return None
    creation_id = data["id"]
    print(f"  -> Đã tạo container. Creation ID: {creation_id}")
    return creation_id


def wait_for_container_ready(container_id, access_token):
    """Chờ container xử lý xong (cần cho VIDEO / CAROUSEL)."""
    url = f"{THREADS_API_BASE}/{container_id}"
    deadline = time.time() + CONTAINER_WAIT_TIMEOUT
    last_status = None
    while time.time() < deadline:
        response = requests.get(
            url,
            params={
                "fields": "status,error_message",
                "access_token": access_token,
            },
        )
        data = response.json()
        status = data.get("status")
        last_status = status
        if status == "FINISHED":
            print(f"  -> Container {container_id} sẵn sàng (FINISHED).")
            return True
        if status in ("ERROR", "EXPIRED"):
            print(f"Lỗi container {container_id}: {data}")
            return False
        print(f"  -> Container {container_id} đang xử lý ({status})...")
        time.sleep(CONTAINER_POLL_INTERVAL)
    print(f"Lỗi: hết thời gian chờ container {container_id} (status={last_status}).")
    return False


def publish_threads_container(user_id, creation_id, access_token):
    url = f"{THREADS_API_BASE}/{user_id}/threads_publish"
    response = requests.post(
        url,
        data={"creation_id": creation_id, "access_token": access_token},
    )
    data = response.json()
    if "id" in data:
        published_id = data["id"]
        print(f"  -> Publish THÀNH CÔNG! Post ID: {published_id}")
        return published_id
    print(f"Lỗi publish: {data}")
    return None


def post_carousel(
    text,
    image_url,
    video_url,
    reply_to_id=None,
    user_id=None,
    access_token=None,
):
    """Đăng 1 bài carousel: ảnh + video (Threads yêu cầu tối thiểu 2 media)."""
    print(f"  -> media_type=CAROUSEL, image={image_url}, video={video_url}")

    image_id = create_threads_container(
        user_id,
        {
            "media_type": "IMAGE",
            "image_url": image_url,
            "is_carousel_item": "true",
            "access_token": access_token,
        },
    )
    if not image_id:
        return None

    video_id = create_threads_container(
        user_id,
        {
            "media_type": "VIDEO",
            "video_url": video_url,
            "is_carousel_item": "true",
            "access_token": access_token,
        },
    )
    if not video_id:
        return None

    if not wait_for_container_ready(image_id, access_token):
        return None
    if not wait_for_container_ready(video_id, access_token):
        return None

    carousel_payload = {
        "media_type": "CAROUSEL",
        "children": f"{image_id},{video_id}",
        "access_token": access_token,
    }
    if text:
        carousel_payload["text"] = text
    if reply_to_id:
        carousel_payload["reply_to_id"] = reply_to_id

    carousel_id = create_threads_container(user_id, carousel_payload)
    if not carousel_id:
        return None
    if not wait_for_container_ready(carousel_id, access_token):
        return None

    return publish_threads_container(user_id, carousel_id, access_token)


def post_to_threads(
    text,
    reply_to_id=None,
    user_id=None,
    access_token=None,
    image_url=None,
    video_url=None,
):
    media_type, image_url, video_url = resolve_media(image_url, video_url)
    text = _clean_str(text)

    if media_type == "TEXT" and not text:
        print("Lỗi: bài TEXT cần có nội dung.")
        return None

    if media_type == "CAROUSEL":
        return post_carousel(
            text,
            image_url,
            video_url,
            reply_to_id=reply_to_id,
            user_id=user_id,
            access_token=access_token,
        )

    payload_create = {
        "media_type": media_type,
        "access_token": access_token,
    }
    if text:
        payload_create["text"] = text
    if media_type == "IMAGE":
        payload_create["image_url"] = image_url
    elif media_type == "VIDEO":
        payload_create["video_url"] = video_url
    if reply_to_id:
        payload_create["reply_to_id"] = reply_to_id

    media_url = image_url or video_url
    print(f"  -> media_type={media_type}" + (f", url={media_url}" if media_url else ""))
    creation_id = create_threads_container(user_id, payload_create)
    if not creation_id:
        return None

    if media_type == "VIDEO":
        if not wait_for_container_ready(creation_id, access_token):
            return None
    else:
        time.sleep(5)

    return publish_threads_container(user_id, creation_id, access_token)


def post_content_chain(df_origin, row_idx, user_id, access_token):
    """Đăng bài chính (+ media) rồi reply bằng sub_content. Trả về True nếu thành công."""
    content = _clean_str(df_origin.loc[row_idx, "content"])
    sub_content = _clean_str(df_origin.loc[row_idx, "sub_content"])
    image_url = df_origin.loc[row_idx, "image_url"]
    video_url = df_origin.loc[row_idx, "video_url"]

    print("content:", content)
    print("sub_content:", sub_content)
    print("Bây giờ là:", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    print("=== ĐANG ĐĂNG BÀI CHÍNH ===")
    post_1_id = post_to_threads(
        content,
        user_id=user_id,
        access_token=access_token,
        image_url=image_url,
        video_url=video_url,
    )
    if not post_1_id:
        print("\n❌ Thất bại ở bài chính, đã dừng chuỗi bài.")
        return False

    if not sub_content:
        print("\n🎉 Hoàn tất! Không có sub_content nên bỏ qua reply.")
        return True

    time.sleep(2)
    print("\n=== ĐANG ĐĂNG REPLY (sub_content) ===")
    post_2_id = post_to_threads(
        sub_content,
        reply_to_id=post_1_id,
        user_id=user_id,
        access_token=access_token,
    )
    if not post_2_id:
        print("\n❌ Bài chính đã lên nhưng reply thất bại.")
        return False

    print("\n🎉 Hoàn tất chuỗi bài (chính + reply).")
    return True


def run_posting_session(file_path, user_id, access_token, session_start, session_end):
    while True:
        df_origin = load_vocab_df(file_path)
        df_filtered = filter_session_posts(df_origin, session_start, session_end)
        posted_count = len(df_filtered)
        print(f"Đã đăng trong phiên: {posted_count}/{CONST_MAX_POST}")

        if not can_post_more(session_end, posted_count):
            break

        pending = df_origin[df_origin["upload_time"].isna()]
        if pending.empty:
            print("⚠️  Hết bài chưa đăng trong file Excel.")
            break

        row_idx = pending.index[0]
        print(f"\n--- Đăng bài #{posted_count + 1} (dòng {row_idx}) ---")
        if post_content_chain(df_origin, row_idx, user_id, access_token):
            uploaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            write_upload_time(file_path, row_idx, uploaded_at)
            df_origin = load_vocab_df(file_path)
            df_filtered = filter_session_posts(df_origin, session_start, session_end)
            print(f"Đã đăng trong phiên: {len(df_filtered)}/{CONST_MAX_POST}")

        if not can_post_more(session_end, len(df_filtered)):
            break

        sleep_between_posts(session_end, file_path)


def run_daily_cycle(file_path, user_id, access_token):
    while True:
        session_start, session_end = wait_until_session_start()
        run_posting_session(
            file_path, user_id, access_token, session_start, session_end
        )

        print(f"\n⏰ Kết thúc phiên lúc {session_end.strftime('%H:%M')}.")
        wait_until_datetime(session_end)

        next_session_start = session_end.replace(
            hour=SESSION_START_HOUR, minute=0, second=0, microsecond=0
        )
        print(
            f"\n🌙 Chờ đến {next_session_start.strftime('%d/%m %H:%M')} "
            "để bắt đầu phiên mới..."
        )
        wait_until_datetime(next_session_start)


if __name__ == "__main__":
    load_dotenv(
        r"C:\Users\vinhdt912\Documents\thuy_onedrive\OneDrive\Tools\Threads\number_la_gi\.env"
    )

    user_id = os.getenv("USER_ID", "me")
    access_token = os.getenv("LONG_LIVED_TOKEN")
    if not access_token or access_token.startswith("YOUR_"):
        raise SystemExit(
            "Thiếu LONG_LIVED_TOKEN trong .env — hãy thay bằng token thật."
        )

    file_path = (
        r"C:\Users\vinhdt912\Documents\thuy_onedrive\OneDrive\Tools\Threads\number_la_gi\vocab_data.xlsx"
    )

    print(load_vocab_df(file_path).head())
    run_daily_cycle(file_path, user_id, access_token)
