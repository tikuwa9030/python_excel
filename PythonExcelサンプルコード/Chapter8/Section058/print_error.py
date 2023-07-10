try:
    result = 1 / 0
except ZeroDivisionError as err:
    print('0で徐算しました', err)
finally:
    print('end')
