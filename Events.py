from dataclasses import dataclass
from typing import List


@dataclass
class event:
    startHour: int
    startMinutes: int
    endHour: int
    endMinutes: int
    startDay: int
    endDay: int
    startMonth: int
    endMonth: int
    year: int
    summary: str


@dataclass
class shiftsSchedule:
    shifts: List[event]




    # @property
    # def summary(self):




