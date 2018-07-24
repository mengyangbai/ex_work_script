class Solution(object):
    def numJewelsInStones(self, J, S):
        """
        :type J: str
        :type S: str
        :rtype: int
        """
        jewset = set(J)
        res = 0
        for c in S:
            if c in jewset:
                res += 1

        return res
