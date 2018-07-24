class Solution(object):
    def spiralOrder(self, matrix):
        """
        :type matrix: List[List[int]]
        :rtype: List[int]
        """
        k, l = 0, 0
        m, n = len(matrix) - 1, len(matrix[0]) - 1

        res = []
        while k <= m and l <= n:

            i = l
            while i <= n:
                res.append(matrix[k][i])
                i = i + 1

            k = k + 1

            i = k
            while i <= m:
                res.append(matrix[i][n])
                i = i + 1

            n = n - 1

            if k <= m:
                i = n
                while i >= l:
                    res.append(matrix[m][i])
                    i = i - 1

            m = m - 1

            if l <= n:
                i = m
                while i >= k:
                    res.append(matrix[i][l])
                    i = i - 1

                l = l + 1

        return res

    def a(self):
        pass


t = Solution()

input = [[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]]

print(t.spiralOrder(input))
