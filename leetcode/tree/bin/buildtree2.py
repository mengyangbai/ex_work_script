# Definition for a binary tree node.
class TreeNode:
    def __init__(self, x):
        self.val = x
        self.left = None
        self.right = None


preorder = [4, 1, 3, 2]
inorder = [1, 2, 3, 4]


def buildTree(preorder, inorder):
    """
    :type preorder: List[int]
    :type inorder: List[int]
    :rtype: TreeNode
    """

    def helper(start, end):
        def search(inorder, arr, start, end, value):
            for i in range(start, end + 1):
                if arr[i] == value:
                    return i

        if start > end:
            return None

        res = TreeNode(preorder[preIndex])
        preIndex += 1

        root_point_index = inorder.index(preorder[start])

        if start == end:
            return res

        root_point_index = search(inorder, start, end, res.val)

        res.left = helper(start, root_point_index - 1)
        res.right = helper(1 + root_point_index, end)

        return res

    preIndex = 0

    return helper(0, len(preorder) - 1)


a = buildTree(preorder, inorder)
