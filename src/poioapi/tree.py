import xlwt

class Tree(object):

    def __init__(self, id=""):
        self.parent = None
        self.children = list()
        self.id = id
        self.tier_id = None
        self.node_info = None
        self.data = ""
        self.depth = 0
        self.marked_as_filtered = True

    def append_child(self, id=""):
        self.children.append(Tree())
        self.children[-1].parent = self
        self.children[-1].id = id
        self.children[-1].depth = self.depth + 1

    def get_child_index(self, id):
        i = 0
        for child in self.children:
            if child.id == id:
                return i
            else:
                i += 1

    def find_node_by_id(self, id):
        if self.children and len(self.children) > 0:
            for child in self.children:
                result = child.find_node_by_id(id)
                if result:
                    return result
        elif self.id == id:
            return self

    def increase_depth(self):
        for child in self.children:
            child.increase_depth()
        self.depth += 1

    # reattach one node - self to another node - goal
    def reattach_node(self, goal):
        if self.parent:
            index = self.parent.get_child_index(self.id)
            self.parent.children.pop(index)
            self.parent = goal
        goal.append_child("")
        goal.children[-1] = self
        goal.children[-1].increase_depth()

    @staticmethod
    def spaces(depth):
        result = ""
        for i in range(0, depth):
            result += "        "
        return result

    def print(self, filtered=False):
        if self.id == "root":
            for child in self.children:
                child.print(filtered)
        elif not filtered or (filtered and self.marked_as_filtered):
            if self.tier_id:
                print(Tree.spaces(self.depth) + "tier_id: " + self.tier_id)
            print(Tree.spaces(self.depth) + "data: " + self.data)
            if self.children:
                for child in self.children:
                    child.print(False)

    def print_to_xls(self, output_file, filtered=False):
        font = xlwt.Font()
        font.name = 'Arial'
        font.colour_index = 0
        font.bold = True

        style = xlwt.XFStyle()
        style.font = font

        wb = xlwt.Workbook()

        tier_col = dict()

        def print_block(node, sheet, offset=0):
            for child in node.children:
                print_block(child, sheet, offset)
            tier_col[node.tier_id] += 1
            sheet.write(tiers.index(node.tier_id)+offset, tier_col[node.tier_id], node.data, style)

        tiers = self.get_tiers()
        for child in self.children:
            if filtered and not child.marked_as_filtered:
                continue

            sheet = wb.add_sheet(child.id)

            child_index = self.children.index(child)
            if child_index > 0:
                left_context = self.children[child_index-1]
            else:
                left_context = None
            if left_context:
                row = 0
                for tier in tiers:
                    sheet.write(row, 0, tier, style)
                    row += 1
                    tier_col[tier] = 1
                print_block(left_context, sheet, 0)

            row = len(tiers) + 2
            for tier in tiers:
                sheet.write(row, 0, tier, style)
                row += 1
                tier_col[tier] = 1
            print_block(child, sheet, len(tiers) + 2)

            if child_index < len(self.children) - 1:
                right_context = self.children[child_index + 1]
            else:
                right_context = None
            if right_context:
                row = 2 * len(tiers) + 4
                for tier in tiers:
                    sheet.write(row, 0, tier, style)
                    row += 1
                    tier_col[tier] = 1
                print_block(right_context, sheet, 2*len(tiers) + 4)

        wb.save(output_file)

    def get_tiers(self):
        tier_list = list()

        def f(node):
            for child in node.children:
                if child.tier_id and (child.tier_id not in tier_list):
                    tier_list.append(child.tier_id)
                f(child)

        f(self)
        return tier_list

    def get_nodes_of_tier(self, tier):
        nodes_list = list()

        def f(node):
            for child in node.children:
                if child.tier_id == tier:
                    nodes_list.append(child)
                f(child)

        f(self)
        return nodes_list

    def node_modify(self):
        for child in self.children:
            if len(child.children) == 0:
                for sibling in child.get_siblings():
                    if sibling.tier_id != child.tier_id:
                        sibling.reattach_node(child)
            child.node_modify()

    # if a child of self has no children attach to it all siblings of other tier ids and call the same for that child
    def tree_modify(self):
        if self.id == "root":
            for child in self.children:
                child.node_modify()
        else:
            self.node_modify()

    # search: [ [str1, str2], [str3], [str4, [str5, str6] ] ]
    # equivalent to (str1 and str2) or str3 or (str4) or (str5 and str6)
    # fields: [tier_id1, tier_id2]

    # marks only those children of root which suit search query
    def filter(self, search, tiers=None):
        if not search:
            return None

        # build search chains from search and compare them with a tree chain
        def search_chain(search, tree_chain):
            if type(search) == list:
                if len(search) == 0:
                    return True
                if len(search) == 1:
                    return search_chain(search[0], tree_chain)
                if len(search) > 1:
                    if type(search[0]) == list:
                        return search_chain(search[0], tree_chain) or search_chain(search[1:], tree_chain)
                    else:
                        return search_chain(search[0], tree_chain) and search_chain(search[1:], tree_chain)
            else:
                result = False
                for element in tree_chain:
                    # + regexp and ==
                    if element.find(search) != -1:
                        result = True
                return result

        # build tree chains and call check function for each of them
        # this variation of function builds chains based on ways from root to leaves
        def tree_chain(node, chain):
            if tiers:
                if node.tier_id in tiers:
                    chain.append(node.data)
            else:
                chain.append(node.data)
            init_depth = node.depth
            if node.children:
                for child in node.children:
                    if tree_chain(child, chain):
                        return True
                    chain = chain[0:init_depth]
                return False
            else:
                # here we have a chain to check if suits the search
                return search_chain(search, chain)

        for child in self.children:
            if not tree_chain(child, list()):
                child.marked_as_filtered = False

    # returns list of children which were marked as filtered
    def return_filtered_children(self):
        result = list()
        for child in self.children:
            if child.marked_as_filtered:
                result.append(child)
        return result

    # resets filters (all marked as filtered = True)
    def reset_filters(self):
        self.marked_as_filtered = True
        for child in self.children:
            child.reset_filters()

    def get_siblings(self):
        siblings = list()
        if self.parent:
            if self.parent.children:
                for child in self.parent.children:
                    if child.id != self.id:
                        siblings.append(child)
        return siblings


