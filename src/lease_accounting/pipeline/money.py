from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import math


COP_QUANTUM = Decimal("1")


def round_cop(value, default=None):
    """Round a monetary value to whole Colombian pesos using half-up."""
    if value is None:
        return default
    try:
        if isinstance(value, float) and math.isnan(value):
            return default
        decimal_value = Decimal(str(value).strip())
        if not decimal_value.is_finite():
            return default
    except (InvalidOperation, TypeError, ValueError):
        return default
    return int(decimal_value.quantize(COP_QUANTUM, rounding=ROUND_HALF_UP))


def allocate_cop(total, weights, residual_index: int | None = None) -> list[int]:
    """Allocate whole pesos proportionally while preserving the rounded total."""
    parsed_weights = []
    for weight in weights:
        try:
            parsed = Decimal(str(weight))
        except (InvalidOperation, TypeError, ValueError):
            parsed = Decimal(0)
        parsed_weights.append(max(parsed, Decimal(0)))

    if not parsed_weights:
        return []

    rounded_total = round_cop(total, 0)
    weight_total = sum(parsed_weights, Decimal(0))
    if weight_total <= 0 or rounded_total == 0:
        return [0 for _ in parsed_weights]

    exact = [Decimal(rounded_total) * weight / weight_total for weight in parsed_weights]
    allocations = [int(value.quantize(COP_QUANTUM, rounding=ROUND_HALF_UP)) for value in exact]
    residual = rounded_total - sum(allocations)
    if residual:
        start = residual_index if residual_index is not None else max(
            range(len(parsed_weights)),
            key=parsed_weights.__getitem__,
        )
        start %= len(allocations)
        direction = 1 if residual > 0 else -1
        for offset in range(abs(residual)):
            allocations[(start + offset) % len(allocations)] += direction
    return allocations
