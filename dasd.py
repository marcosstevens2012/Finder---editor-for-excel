#coding=utf-8


def my_func(r, n):
  for i in xrange(n): r = hashlib.sha1(r[:9]).hexdigest()
  return r

#calcular el valor de:
print my_func("0123456789012345678901234567890123456789", 9999999999999999)

