import time
from functools import lru_cache

# ۱. تابع معمولی (بدون بهینه‌سازی)
def fibo_slow(n):
    if n <= 1:
        return n
    return fibo_slow(n - 1) + fibo_slow(n - 2)

# ۲. تابع با استفاده از Memoization
@lru_cache(maxsize=None)
def fibo_fast(n):
    if n <= 1:
        return n
    return fibo_fast(n - 1) + fibo_fast(n - 2)

# تست سرعت تابع کند
print("--- در حال محاسبه با روش معمولی... ---")
start = time.time()
result = fibo_slow(35)
end = time.time()
print(f"نتیجه: {result} | زمان اجرا: {end - start:.4f} ثانیه")

# تست سرعت تابع سریع
print("\n--- در حال محاسبه با روش Memoization... ---")
start = time.time()
result = fibo_fast(35)
end = time.time()
print(f"نتیجه: {result} | زمان اجرا: {end - start:.8f} ثانیه")