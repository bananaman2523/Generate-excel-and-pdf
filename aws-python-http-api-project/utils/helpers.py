from dataclasses import dataclass, field
from typing import List, Dict
@dataclass
class ArrayHolder:
    values: List[int] = field(default_factory=list)

    def add_value(self, value: int):
        self.values.append(value)

    def clear_array(self):
        self.values.clear()

    def __repr__(self):
        return f"ArrayHolder(values={self.values})"


    
@dataclass
class SheetHolder:
    values: Dict[str, int] = field(default_factory=dict)

    def add_value(self, value: str):
        if value in self.values:
            self.values[value] += 1
        else:
            self.values[value] = 1

    def get_count(self, value: str) -> int:
        return self.values.get(value, 0)

    def __repr__(self):
        return f"SheetHolder(values={self.values})"