class Solution(object):
    def generate(self, numRows):
        """
        :type numRows: int
        :rtype: List[List[int]]
        """
        res = []
        i = 1
        while i <= numRows:
            single_row = []
            j = 0
            while j < i:
                if j == 0 or j == i - 1:
                    single_row.append(1)
                else:
                    single_row.append(res[-1][j - 1] + res[-1][j])

                j = j + 1

            res.append(single_row)
            i = i + 1

        return res


a = Solution()

a.generate(5)
