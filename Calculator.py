import xlwings as xw

def add(x, y):
    return x + y

def multiply(x, y):
    return x * y

def calculate(x, y, operation):
    if operation == "SUM":
        return add(x, y)
    elif operation == "PRODUCT":
        return multiply(x, y)
    elif operation == "SUBTRACT":
        return x - y
    elif operation == "DIVIDE":
        if y == 0.0:
            raise ValueError("Division by zero!")
        return x / y
    elif operation == "POW":
        return x ** y
    elif operation == "MOD":
        if y == 0:
            raise ValueError("Modulo by zero!")
        return x % y

def compute():
    wb = xw.Book.caller()
    sheet = wb.sheets["Calculator"]
    x = sheet.range("B2").value
    y = sheet.range("B3").value
    operation = sheet.range("B4").value.strip().upper()
    try:
        result = calculate(x, y, operation)
        sheet.range("B6").value = result
        sheet.range("B7").value = "OK"
    except Exception as e:
        sheet.range("B6").value = ""
        sheet.range("B7").value = str(e)

def main():
    compute()