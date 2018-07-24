class Solution(object):
    def dominantIndex(self, nums):
        """
        :type nums: List[int]
        :rtype: int
        """
        a = []
        a.append((-9999, -1))
        a.append((-9999, -1))
        for i, n in enumerate(nums):
            if n > a[0][0]:
                a[1] = a[0]
                a[0] = (n, i)
            elif n > a[1][0]:
                a[1] = (n, i)

        if a[0][0] < a[1][0] * 2:
            return -1

        return a[0][1]


a = Solution()

nums = [0, 0, 0, 1]
a.dominantIndex(nums)
