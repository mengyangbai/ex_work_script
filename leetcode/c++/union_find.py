class Solution:
    def validTree(self, n, edges):
        """
        :type n: int
        :type edges: List[List[int]]
        :rtype: bool
        """

        def find(i):
            while self.root[i] != -1:
                i = self.root[i]
            return i

        self.root = [-1] * n

        for pair in edges:
            x = find(pair[0])
            y = find(pair[1])
            if (x == y):
                return False
            self.root[x] = y

        return len(edges) == n - 1


a = Solution()
# print(a.validTree(5, [[0, 1], [0, 2], [0, 3], [1, 4]]))
print(a.validTree(5, [[0, 1], [1, 2], [2, 3], [1, 3], [1, 4]]))
