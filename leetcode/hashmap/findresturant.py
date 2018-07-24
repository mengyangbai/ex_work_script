class Solution(object):
    def findRestaurant(self, list1, list2):
        """
        :type list1: List[str]
        :type list2: List[str]
        :rtype: List[str]
        """
        dic = {}
        for i in range(len(list1)):
            if list1[i] in dic:
                dic[list1[i]][0] = i
            else:
                dic[list1[i]] = [i, 2002]

        for i in range(len(list2)):
            if list2[i] in dic:
                dic[list2[i]][1] = i
            else:
                dic[list2[i]] = [2002, i]

        res = []
        minvalue = 9999
        for key, value in dic.items():
            if value[0] + value[1] == minvalue:
                minvalue = value[0] + value[1]
                res.append(key)
            elif value[0] + value[1] < minvalue:
                res.clear()
                minvalue = value[0] + value[1]
                res.append(key)

        return res
