def printDebug(*args, **kwargs):
    print(f'Args => {args}')
    print(f'Args => {kwargs}')


printDebug(a, b, c=10)
printDebug(c, d=5, x=10)

# input string
# Dict => {'A': 2, 'N': 2, 'X': 1}