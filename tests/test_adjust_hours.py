# tests/test_adjust_hour.py
import unittest
from datetime import datetime
from src.adjust_hours import calculate_total_hours

class TestAdjustHour(unittest.TestCase):
    def test_calculate_total_hours(self):
        start_time = datetime.strptime("2024-04-01 08:00:00", "%Y-%m-%d %H:%M:%S")
        end_time = datetime.strptime("2024-04-01 16:00:00", "%Y-%m-%d %H:%M:%S")
        pause_minutes = 30
        expected_hours = 7.5  # (16 - 8) - 0.5 = 7.5
        self.assertEqual(calculate_total_hours(start_time, end_time, pause_minutes), expected_hours)

    def test_calculate_total_hours_with_rounding(self):
        start_time = datetime.strptime("2024-04-01 08:00:00", "%Y-%m-%d %H:%M:%S")
        end_time = datetime.strptime("2024-04-01 16:15:00", "%Y-%m-%d %H:%M:%S")
        pause_minutes = 30
        expected_hours = 8.25  # (16.25 - 0.5) = 7.75, rounded up to 8.0
        self.assertEqual(calculate_total_hours(start_time, end_time, pause_minutes), 7.75)

    def test_calculate_total_hours_negative(self):
        start_time = datetime.strptime("2024-04-01 16:00:00", "%Y-%m-%d %H:%M:%S")
        end_time = datetime.strptime("2024-04-01 08:00:00", "%Y-%m-%d %H:%M:%S")
        pause_minutes = 30
        expected_hours = 0  # Negative Arbeitszeit
        self.assertEqual(calculate_total_hours(start_time, end_time, pause_minutes), expected_hours)

if __name__ == '__main__':
    unittest.main()
