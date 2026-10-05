"""Small, non-executing formula evaluator. Unsupported Excel functions are explicit."""
from functools import lru_cache
import re

from openpyxl.utils import range_boundaries, get_column_letter

TOKEN = re.compile(r'\s*("(?:[^"]|"")*"|\'(?:[^\']|\'\')+\'![A-Z]+\d+(?::[A-Z]+\d+)?|[A-Za-z_][A-Za-z_ 0-9]*![A-Z]+\d+(?::[A-Z]+\d+)?|[A-Z]+\d+(?::[A-Z]+\d+)?|[A-Z]+|\d+(?:\.\d+)?|<>|<=|>=|[(),=<>+\-*/])')


class Parser:
    def __init__(self, formula):
        text = formula[1:].replace("$", "")
        self.tokens = []
        while text:
            match = TOKEN.match(text)
            if not match:
                raise ValueError(f"Unsupported formula syntax: {text}")
            self.tokens.append(match[1])
            text = text[match.end():]
        self.index = 0

    def peek(self):
        return self.tokens[self.index] if self.index < len(self.tokens) else None

    def pop(self):
        value = self.peek()
        if value is None:
            raise ValueError("Incomplete formula")
        self.index += 1
        return value

    def require(self, token):
        if self.pop() != token:
            raise ValueError(f"Expected {token}")

    def expression(self, minimum=0):
        token = self.pop()
        if token in ("+", "-"):
            left = ("unary", token, self.expression(4))
        elif token == "(":
            left = self.expression()
            self.require(")")
        elif token.startswith('"'):
            left = ("literal", token[1:-1].replace('""', '"'))
        elif re.fullmatch(r"\d+(?:\.\d+)?", token):
            left = ("literal", float(token))
        elif self.peek() == "(":
            self.pop()
            args = []
            while self.peek() != ")":
                args.append(self.expression())
                if self.peek() != ",":
                    break
                self.pop()
            self.require(")")
            left = ("call", token, args)
        else:
            left = ("reference", token)
        precedence = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
                      "+": 2, "-": 2, "*": 3, "/": 3}
        while self.peek() in precedence and precedence[self.peek()] >= minimum:
            op = self.pop()
            left = ("binary", op, left, self.expression(precedence[op] + 1))
        return left

    def parse(self):
        result = self.expression()
        if self.index != len(self.tokens):
            raise ValueError("Unconsumed formula tokens")
        return result


def flatten(values):
    for value in values:
        if isinstance(value, list):
            yield from flatten(value)
        else:
            yield value


def evaluator(workbook):
    visiting = set()

    @lru_cache(None)
    def cell(sheet, address):
        key = (sheet, address)
        if key in visiting:
            raise ValueError("Circular formula reference")
        value = workbook[sheet][address].value
        if value is None:
            return ""
        if isinstance(value, str) and value.startswith("="):
            visiting.add(key)
            try:
                return evaluate(Parser(value).parse(), sheet)
            finally:
                visiting.remove(key)
        return value

    def evaluate(node, sheet):
        kind = node[0]
        if kind == "literal":
            return node[1]
        if kind == "reference":
            address = node[1]
            if address in ("TRUE", "FALSE"):
                return address == "TRUE"
            if "!" in address:
                sheet, address = address.split("!", 1)
                sheet = sheet.strip("'").replace("''", "'")
            if ":" not in address:
                return cell(sheet, address)
            c1, r1, c2, r2 = range_boundaries(address)
            return [cell(sheet, f"{get_column_letter(c)}{r}")
                    for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
        if kind == "unary":
            value = evaluate(node[2], sheet)
            return value if node[1] == "+" else -value
        if kind == "binary":
            op, a, b = node[1], evaluate(node[2], sheet), evaluate(node[3], sheet)
            operations = {"=": lambda: a == b, "<>": lambda: a != b,
                          "<": lambda: a < b, ">": lambda: a > b,
                          "<=": lambda: a <= b, ">=": lambda: a >= b,
                          "+": lambda: a + b, "-": lambda: a - b,
                          "*": lambda: a * b, "/": lambda: a / b}
            return operations[op]()
        name, args = node[1], node[2]
        if name == "IF":
            if len(args) != 3:
                raise ValueError("IF needs three arguments")
            return evaluate(args[1] if evaluate(args[0], sheet) else args[2], sheet)
        values = list(flatten(evaluate(arg, sheet) for arg in args))
        if name == "OR":
            return any(values)
        if name == "AND":
            return all(values)
        if name == "SUM":
            return sum(v for v in values if isinstance(v, (int, float)))
        raise ValueError(f"Unsupported formula function: {name}")

    return cell
