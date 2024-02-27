class Achievement:
    def __init__(self, name: str, title: str, max_value: int, value: int = 0) -> None:
        self.name = name
        self.title = title
        self.value = value
        self.max_value = max_value