"""Small, deterministic Excel value readers; never execute workbook code."""
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import math
import re

from openpyxl.utils.datetime import from_excel


def excel_date(value, epoch, parse):
    """Use the workbook's epoch for unformatted serial dates, not a guessed epoch."""
    parsed = parse(value)
    if parsed is not None:
        return parsed
    if isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
        try:
            converted = from_excel(value, epoch)
            if isinstance(converted, datetime) and 1950 <= converted.year <= 2100:
                return converted.date()
        except (OverflowError, ValueError):
            pass
    return None


def weight_number(value, parse):
    """Accept an explicit kg suffix only; grams/tonnes are not silently reinterpreted."""
    parsed = parse(value)
    if parsed is not None:
        return parsed
    if isinstance(value, str):
        match = re.fullmatch(r'\s*(.+?)\s*(?:kg|kgs|kilograms?)\s*', value, re.I)
        if match:
            return parse(match[1])
    return None


def quantity_number(value, parse):
    """Read the register's explicit lot/pcs quantity labels without guessing."""
    parsed = parse(value)
    if parsed is not None:
        return parsed
    if isinstance(value, str):
        match = re.fullmatch(r'\s*(.+?)\s+(?:lots?|pcs|pieces?|units?)\s*', value, re.I)
        if match:
            return parse(match[1])
    return None


def formula_number(formula, references):
    """Recalculate bounded arithmetic, SUM and ROUND from approved same-row inputs.

    No cached values, external links, arbitrary Excel functions, ranges or Python
    eval. The caller decides which cell references can be used.
    """
    if not isinstance(formula, str) or not formula.startswith('=') or len(formula) > 500:
        return None
    expression = re.sub(r'\s+|\$', '', formula[1:]).upper()
    pattern = re.compile(r'(?:\d+(?:\.\d*)?|\.\d+)(?:E[+-]?\d+)?|[A-Z]+\d+|[A-Z]+|[()+*/,-]')
    tokens = pattern.findall(expression)
    if ''.join(tokens) != expression or not tokens or len(tokens) > 100:
        return None
    position = 0

    def take(expected=None):
        nonlocal position
        if position >= len(tokens) or (expected is not None and tokens[position] != expected):
            raise ValueError('Incomplete formula')
        token = tokens[position]
        position += 1
        return token

    def peek():
        return tokens[position] if position < len(tokens) else ''

    def atom(depth):
        if depth > 12:
            raise ValueError('Formula too deep')
        token = take()
        if token in {'+', '-'}:
            return atom(depth + 1) * (-1 if token == '-' else 1)
        if token == '(':
            result = expression_value(depth + 1)
            take(')')
            return result
        if re.fullmatch(r'[A-Z]+\d+', token):
            value = references.get(token)
            if value is None or not math.isfinite(value):
                raise ValueError('Unapproved reference')
            return float(value)
        if token in {'SUM', 'ROUND'}:
            take('(')
            arguments = [expression_value(depth + 1)]
            while peek() == ',':
                take(',')
                arguments.append(expression_value(depth + 1))
                if len(arguments) > 20:
                    raise ValueError('Too many arguments')
            take(')')
            if token == 'SUM':
                return sum(arguments)
            if len(arguments) != 2 or not arguments[1].is_integer() or not -8 <= arguments[1] <= 8:
                raise ValueError('Unsupported rounding')
            quantum = Decimal(1).scaleb(-int(arguments[1]))
            return float(Decimal(str(arguments[0])).quantize(quantum, rounding=ROUND_HALF_UP))
        return float(token)

    def product(depth):
        result = atom(depth)
        while peek() in {'*', '/'}:
            operator = take()
            right = atom(depth)
            result = result * right if operator == '*' else result / right
        return result

    def expression_value(depth):
        result = product(depth)
        while peek() in {'+', '-'}:
            operator = take()
            right = product(depth)
            result = result + right if operator == '+' else result - right
        return result

    try:
        result = expression_value(0)
        return result if position == len(tokens) and math.isfinite(result) else None
    except (ValueError, OverflowError, ZeroDivisionError, InvalidOperation, TypeError):
        return None
