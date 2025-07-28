"""
時間工具模組
負責產生每10分鐘的時間區段和時間相關功能
"""

from datetime import datetime, timedelta
import config

def generate_time_segments(start_time=None, end_time=None, interval_minutes=None):
    """
    產生時間區段列表
    
    Args:
        start_time (str): 開始時間 (HH:MM 格式)，None 則使用 config 設定
        end_time (str): 結束時間 (HH:MM 格式)，None 則使用 config 設定
        interval_minutes (int): 時間間隔分鐘數，None 則使用 config 設定
        
    Returns:
        list: 時間區段列表，每個元素為 (start_time, end_time) 元組
    """
    # 使用預設值
    start_time = start_time or config.START_TIME
    end_time = end_time or config.END_TIME
    interval_minutes = interval_minutes or config.TIME_INTERVAL_MINUTES
    
    # 解析時間
    start_dt = datetime.strptime(start_time, "%H:%M")
    end_dt = datetime.strptime(end_time, "%H:%M")
    
    # 如果結束時間小於開始時間，表示跨日
    if end_dt < start_dt:
        end_dt += timedelta(days=1)
    
    # 產生時間區段
    time_segments = []
    current_dt = start_dt
    
    while current_dt < end_dt:
        # 計算下一個時間點
        next_dt = current_dt + timedelta(minutes=interval_minutes)
        
        # 如果超過結束時間，調整為結束時間
        if next_dt > end_dt:
            next_dt = end_dt
        
        # 格式化時間
        segment_start = current_dt.strftime("%H:%M")
        segment_end = next_dt.strftime("%H:%M")
        
        time_segments.append((segment_start, segment_end))
        
        # 移動到下一個區段
        current_dt = next_dt
    
    return time_segments

def generate_daily_time_segments(date=None):
    """
    產生指定日期的完整時間區段
    
    Args:
        date (str): 日期 (YYYY-MM-DD 格式)，None 則使用今天
        
    Returns:
        list: 時間區段列表，每個元素為 (start_time, end_time) 元組
    """
    if date is None:
        date = datetime.now().strftime("%Y-%m-%d")
    
    print(f"產生 {date} 的時間區段...")
    segments = generate_time_segments()
    print(f"共產生 {len(segments)} 個時間區段")
    
    return segments

def format_time_for_display(time_str):
    """
    格式化時間顯示
    
    Args:
        time_str (str): 時間字串 (HH:MM 格式)
        
    Returns:
        str: 格式化後的時間字串
    """
    try:
        dt = datetime.strptime(time_str, "%H:%M")
        return dt.strftime("%H:%M")
    except ValueError:
        return time_str

def get_current_time_segment():
    """
    取得當前時間所屬的時間區段
    
    Returns:
        tuple: (start_time, end_time) 或 None
    """
    now = datetime.now()
    current_time = now.strftime("%H:%M")
    
    # 取得今天的時間區段
    segments = generate_daily_time_segments()
    
    # 尋找當前時間所屬的區段
    for start_time, end_time in segments:
        if start_time <= current_time <= end_time:
            return (start_time, end_time)
    
    return None

def is_time_in_segment(time_str, start_time, end_time):
    """
    檢查時間是否在指定區段內
    
    Args:
        time_str (str): 要檢查的時間 (HH:MM 格式)
        start_time (str): 區段開始時間 (HH:MM 格式)
        end_time (str): 區段結束時間 (HH:MM 格式)
        
    Returns:
        bool: 是否在區段內
    """
    try:
        time_dt = datetime.strptime(time_str, "%H:%M")
        start_dt = datetime.strptime(start_time, "%H:%M")
        end_dt = datetime.strptime(end_time, "%H:%M")
        
        # 處理跨日情況
        if end_dt < start_dt:
            end_dt += timedelta(days=1)
            if time_dt < start_dt:
                time_dt += timedelta(days=1)
        
        return start_dt <= time_dt <= end_dt
        
    except ValueError:
        return False

def get_segment_index(start_time, end_time):
    """
    取得時間區段在一天中的索引
    
    Args:
        start_time (str): 開始時間 (HH:MM 格式)
        end_time (str): 結束時間 (HH:MM 格式)
        
    Returns:
        int: 區段索引 (0-based)，找不到則返回 -1
    """
    segments = generate_daily_time_segments()
    
    for i, (seg_start, seg_end) in enumerate(segments):
        if seg_start == start_time and seg_end == end_time:
            return i
    
    return -1

def validate_time_format(time_str):
    """
    驗證時間格式
    
    Args:
        time_str (str): 時間字串
        
    Returns:
        bool: 格式是否正確
    """
    try:
        datetime.strptime(time_str, "%H:%M")
        return True
    except ValueError:
        return False

def get_time_difference(time1, time2):
    """
    計算兩個時間的差異（分鐘）
    
    Args:
        time1 (str): 時間1 (HH:MM 格式)
        time2 (str): 時間2 (HH:MM 格式)
        
    Returns:
        int: 差異分鐘數
    """
    try:
        dt1 = datetime.strptime(time1, "%H:%M")
        dt2 = datetime.strptime(time2, "%H:%M")
        
        # 計算差異
        diff = abs((dt2 - dt1).total_seconds() / 60)
        return int(diff)
        
    except ValueError:
        return 0 