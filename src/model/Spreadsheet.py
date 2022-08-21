import re


class Sheet:

    def __init__(self) -> None:
        self.excel = {}

    def get(self, column: str):
        """
        该函数的目的是获取在column列的值，内容为等式的话需要计算并返回其结果。举例，若在"A1"列存储的值为"=7+3"，sheet.get("A1")应返回"10"。
        :param column: 列数
        :return: 该列存储的值，默认为空字符串
        """
        return self.solve(self.excel.get(column, ""))

    def get_literal(self, column: str):
        """
        该函数的目的是获取在column列的字符串值，内容为等式的话不需要计算，直接返回字符串。举例，若在"A1"列存储的值为"=7+3"，sheet.getLiteral("A1")应返回"=7+3"。
        :param column: 列数
        :return: 该列存储的字符串值，默认为空字符串
        """
        return self.excel.get(column, "")

    def put(self, column: str, value: str):
        """
        该函数的目的是在column列存储value的值。如果该列已经被占用，则替换为当前值。
        :param column: 列数
        :param value: 在该列需要存储的值
        """
        self.excel[column] = value
        return value

    def solve(self, value: str):
        """
        该函数用于处理value的返回值
        :param value: 列中存储的值
        """
        value = self._solve_number(value)
        value = self._solve_operation(value)
        return value

    def _solve_number(self, value):
        """
        当value是数字时去掉空格
        :param value: 列中存储的值
        """
        try:
            int(value.strip())
        except ValueError:
            return value
        else:
            return value.strip()

    def _solve_operation(self, value):
        """
        计算表达式的值
        :param value: 列中存储的值
        """
        if value and value[0] == "=":
            cal = Calculator(self)
            return cal.calculate(value[1:]) or ''
        return value


class Calculator:
    repeat = None

    def __init__(self, sheet):
        self.sheet = sheet

    def formats(self, s):
        """
        格式化表达式
        :param s: 表达式
        """
        s = s.replace(' ', '')
        return s

    def mul_div(self, s):
        """
        乘除计算
        :param s: 表达式
        """
        r = re.compile(r'[\w\.]+[\*/]-?[\w\.]+')
        while re.search(r'[\*/]', s):
            ma = re.search(r, s).group()
            li = re.findall(r'(-?[\w\.]+|\*|/)', ma)
            if li[1] == '*':
                result = str(
                    get_value(self.sheet, li[0], self) * get_value(self.sheet, li[2], self))
            else:
                result = str(
                    int(get_value(self.sheet, li[0], self) / get_value(self.sheet, li[2], self)))
            s = s.replace(ma, result, 1)
        return s

    def add_sub(self, s):
        """ 
        加减计算
        :param s: 表达式
        """
        if '(' in s or ')' in s:
            raise ValueError('#Error')
        li = re.findall(r'([\w\.]+|\+|-)', s)
        nums = 0
        # 如果没有加号和减号
        if len(li) == 0:
            return str(get_value(self.sheet, s, self))
        for i in range(len(li)):
            if li[i] == '-':
                li[i] = '+'
                li[i + 1] = get_value(self.sheet, li[i + 1], self) * -1
        for i in li:
            if i == '+':
                i = 0
            nums = nums + get_value(self.sheet, i, self)
        return str(nums)

    def simple(self, x):
        """
        处理不带括号的表达式
        :param x: 表达式
        """
        return self.add_sub(self.mul_div(x))

    def calculate(self, x):
        """
        计算一个复杂表达式的值
        :param x: 复杂表达式
        """
        x = self.formats(x)
        try:
            while '(' in x and ')' in x:
                reg = re.compile(r'\([^\(\)]+\)')
                ma = re.search(reg, x).group()
                result = self.simple(ma[1:-1])
                x = x.replace(ma, result, 1)
            return self.simple(x)
        except RuntimeError:
            return '#Circular'
        except Exception as e:
            return '#Error'


def is_number(value):
    """
    判断字符串是不是数字
    :param value: 要判断的字符串
    """
    try:
        int(value)
        return True
    except Exception:
        return False


def get_value(sheet, value, cal):
    """
    获取表达式或者列的值
    :param sheet: Sheet的一个实例
    :param value: 列号或者表达式或者数字字符串
    :param cal: Calculator的一个实例
    """
    # 如果是数字，直接返回
    if is_number(value):
        return int(value)
    origin_value = value
    # 不是数字就引用了其他列的值
    flag = False
    if cal.repeat is None:
        flag = True
        cal.repeat = {}
    cal.repeat[origin_value] = 1  # 1表示正在获取value的值
    value = sheet.excel.get(value)
    if value and not is_number(value):
        # value是一个表达式
        if value[0] != '=':
            raise ValueError('#Error')
        value = value[1:]
        if cal.repeat.get(value) == 1:
            raise RuntimeError('#Circular')
        value = cal.calculate(value)
    if value:
        cal.repeat[origin_value] = 2  # 2表示获取成功
    if flag is True:
        cal.repeat = None
    return int(value or 0)
