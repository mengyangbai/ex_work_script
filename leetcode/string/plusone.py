class Solution(object):
    def plusOne(self, digits):
        """
        :type digits: List[int]
        :rtype: List[int]
        """
        number = int("".join(str(x) for x in digits)) + 1
        res = [int(x) for x in str(number)]
        return res


a = Solution()
a.plusOne([1, 2, 3])
