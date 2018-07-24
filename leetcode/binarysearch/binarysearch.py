class Solution(object):
    def findClosestElements(self, arr, k, x):
        """
        :type arr: List[int]
        :type k: int
        :type x: int
        :rtype: List[int]
        """
        start = 0
        end = len(arr) - 1
        while start < end:
            mid = (start + end) // 2
            if arr[mid] == x:
                start = mid
                while arr[start] == x:
                    start = start - 1
                end = start + 1
            elif arr[mid] < x:
                start = mid
            else:
                end = mid

        return arr[start:start + k]
