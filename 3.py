str = 'Welcome to Python Examples'


chunks = [str[i:i+1] for i in range(0, len(str), 1)]
print(chunks[1])
