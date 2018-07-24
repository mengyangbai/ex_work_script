class Solution:
    def kEmptySlots(self, flowers, k):

        maxn = max(flowers)
        nums = [0] * (maxn + 1)
        ft = FenwickTree(maxn)
        for i, v in enumerate(flowers):
            ft.add(v, 1)
            nums[v] = 1
            if v >= k and ft.sum(v) - ft.sum(v - k - 2) == 2 and nums[v - k -
                                                                      1]:
                return i + 1
            if v + k + 1 <= maxn and ft.sum(v + k + 1) - ft.sum(
                    v - 1) == 2 and nums[v + k + 1]:
                return i + 1
        return -1


class FenwickTree(object):
    def __init__(self, n):
        self.n = n
        self.sums = [0] * (n + 1)

    def add(self, x, val):
        while x <= self.n:
            self.sums[x] += val
            x += self.lowbit(x)

    def lowbit(self, x):
        return x & -x

    def sum(self, x):
        res = 0
        while x > 0:
            res += self.sums[x]
            x -= self.lowbit(x)
        return res


a = Solution()
a.kEmptySlots([1, 3, 2], 1)
