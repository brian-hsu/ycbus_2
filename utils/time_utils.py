from datetime import datetime, time
from typing import Tuple

class TimeHandler:
    @staticmethod
    def parse_time(time_str: str) -> Tuple[int, int]:
        """解析時間字串為小時和分鐘"""
        try:
            hour, minute = map(int, time_str.split(':'))
            return hour, minute
        except ValueError as e:
            raise ValueError(f"Invalid time format: {time_str}") from e

    @staticmethod
    def adjust_time(hour: int, minute: int, adjustment: int) -> str:
        """調整時間"""
        total_minutes = hour * 60 + minute + adjustment
        new_hour = (total_minutes // 60) % 24
        new_minute = total_minutes % 60
        return f"{new_hour:02d}:{new_minute:02d}" 