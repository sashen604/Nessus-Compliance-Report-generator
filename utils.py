# utils.py

from collections import Counter


def count_l1_l2(data):
    """
    Count L1, L2 and total checks.
    Supports:
    - summary rows (tuple)
    - detailed items (dict)
    """

    levels = []

    for item in data:
        if isinstance(item, tuple):  # summary
            levels.append(item[1])
        elif isinstance(item, dict):  # detailed
            levels.append(item.get("level", ""))

    counter = Counter(levels)

    l1 = counter.get("L1", 0)
    l2 = counter.get("L2", 0)

    return l1, l2, l1 + l2