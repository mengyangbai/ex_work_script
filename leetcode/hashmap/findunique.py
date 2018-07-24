class Solution(object):
    """practiswork

    Arguments:
        object {[type]} -- [description]

    Returns:
        [type] -- [description]
    """

    def first_uniqchar(self, s):
        """
        :type s: str
        :rtype: int
        """
        dic = {}
        for c in s:
            if c in dic:
                dic[c] = True
            else:
                dic[c] = False

        for i, v in enumerate(s):
            if not dic[v]:
                return i

        return -1
