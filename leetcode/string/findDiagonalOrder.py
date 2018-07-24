import collections


class Solution(object):
    def findDiagonalOrder(self, matrix):
        """
        :type matrix: List[List[int]]
        :rtype: List[int]
        """
        if len(matrix) == 0: return []
        result = []
        count = collections.defaultdict(list)
        m, n = len(matrix), len(matrix[0])
        for i in range(len(m)):
            for j in range(len(n)):
                count[i + j].append(matrix[i][j])

        for i in range(len(m + n - 1)):
            result += count[i][::-1] if i & 1 == 0 else count[i]

        return result