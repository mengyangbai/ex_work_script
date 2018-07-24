# Definition for a binary tree node.
'''
    Algorithm test
    @author Bain.bai
'''


class TreeNode:
    def __init__(self, x):
        self.val = x
        self.left = None
        self.right = None


class TreeNodeWithCol:
    def __init__(self, node, col):
        self.treenode = node
        self.col = col


class Solution:
    def verticalOrder(self, root):
        """
        :type root: TreeNode
        :rtype: List[List[int]]
        """
        res = []
        if not root:
            return res

        queue = []
        dic = {}
        queue.append(TreeNodeWithCol(root, 0))
        cur_level = 1
        next_level = 0
        min_val = 0
        max_val = 0
        while queue:
            cur_node = queue.pop(0)
            if cur_node.col in dic:
                dic[cur_node.col].append(cur_node.treenode.val)
            else:
                dic[cur_node.col] = [cur_node.treenode.val]

            cur_level -= 1

            if cur_node.treenode.left:
                queue.append(
                    TreeNodeWithCol(cur_node.treenode.left, cur_node.col - 1))
                next_level += 1
                min_val = min(cur_node.col - 1, min_val)

            if cur_node.treenode.right:
                queue.append(
                    TreeNodeWithCol(cur_node.treenode.right, cur_node.col + 1))
                next_level += 1
                max_val = max(cur_node.col + 1, max_val)

            if cur_level == 0:
                cur_level = next_level
                next_level = 0

        for i in range(min_val, max_val + 1):
            res.append(dic[i])

        return res


a = Solution()
node3 = TreeNode(3)
node9 = TreeNode(9)
node20 = TreeNode(20)
node15 = TreeNode(15)
node7 = TreeNode(7)
node3.left = node9
node3.right = node20
node20.left = node15
node20.right = node7

a.verticalOrder(node3)
