class NumMatrix(object):
    def __init__(self, matrix):
        """
        :type matrix: List[List[int]]
        """
        if not matrix or not matrix[0]:
            return

        self.matrix = matrix
        self.m, self.n = len(matrix), len(matrix[0])

        self.tree = [[0] * (self.n + 1) for _ in range(self.m + 1)]
        self.nums = [[0] * self.n for _ in range(self.m)]

        for i in range(0, self.m):
            for j in range(0, self.n):
                self.update(i, j, matrix[i][j])

    def update(self, row, col, val):
        """
        :type row: int
        :type col: int
        :type val: int
        :rtype: void
        """
        if self.m == 0 or self.n == 0:
            return

        delta = val - self.nums[row][col]
        self.nums[row][col] = val
        i = row + 1
        while i <= self.m:
            j = col + 1
            while j <= self.n:
                self.tree[i][j] += delta
                j += j & (-j)
            i += i & (-i)

    def sumRegion(self, row1, col1, row2, col2):
        """
        :type row1: int
        :type col1: int
        :type row2: int
        :type col2: int
        :rtype: int
        """
        if self.m == 0 or self.n == 0:
            return

        return self.sum(row2 + 1, col2 + 1) + self.sum(row1, col1) - self.sum(
            row1, col2 + 1) - self.sum(row2 + 1, col1)

    def sum(self, row, col):
        res = 0
        i = row
        while i > 0:
            j = col
            while j > 0:
                res += self.tree[i][j]
                j -= j & (-j)
            i -= i & (-i)
        return res


a = NumMatrix([[3, 0, 1, 4, 2], [5, 6, 3, 2, 1], [1, 2, 0, 1, 5],
               [4, 1, 0, 1, 7], [1, 0, 3, 0, 5]])

a.sumRegion(2, 1, 4, 3)
a.update(3, 2, 2)
a.sumRegion(2, 1, 4, 3)