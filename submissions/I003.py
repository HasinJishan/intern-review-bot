def primes_upto(n):
    for i in range(1, n):
        for j in range(2, i):
            if i % j == 0:
                break
        else:
            print(i)
