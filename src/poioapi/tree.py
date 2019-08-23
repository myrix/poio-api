import xlwt
import xlrd
from xlrd import XLRDError
from xlutils.copy import copy
import re


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

    def print_to_xls(self, output_file, name="sheet", filtered=False):
        font_main = xlwt.Font()
        font_main.name = 'Arial'
        font_main.colour_index = 0
        font_main.bold = False

        font_tiers = xlwt.Font()
        font_tiers.name = 'Arial'
        font_tiers.colour_index = 0
        font_tiers.bold = True

        font_sep = xlwt.Font()
        font_sep.name = 'Arial'
        font_sep.colour_index = 2
        font_sep.bold = True

        style = xlwt.XFStyle()
        style.font = font_tiers

        style_tiers = xlwt.XFStyle()
        style_tiers.font = font_main

        style_sep = xlwt.XFStyle()
        style_sep.borders.bottom = 6
        style_sep.borders.top = 6
        style_sep.borders.bottom_colour = 2
        style_sep.borders.top_colour = 2

        style_context_sep = xlwt.XFStyle()
        style_context_sep.font = font_sep
        style_context_sep.borders.top = 2
        style_context_sep.borders.top_colour = 3
        style_context_sep.borders.bottom = 2
        style_context_sep.borders.bottom_colour = 3

        wb = xlwt.Workbook()
        flag = False
        try:
            rb = xlrd.open_workbook(output_file, on_demand=True, formatting_info=True)
        except XLRDError as e:
            if str(e) == "File size is 0 bytes":
                flag = True
            else:
                raise ("Got a trouble with xls file opening")
        if not flag:
            wb = copy(rb)

        sheet = wb.add_sheet(name)
        # above: opened output file for editting and added new sheet to it

        tier_col = dict()

        empty = True

        def print_block(node, sheet, row_offset=0):
            for child in node.children:
                print_block(child, sheet, row_offset)
            tier_col[node.tier_id] += 1
            sheet.write(tiers.index(node.tier_id) + row_offset, tier_col[node.tier_id], node.data, style)

        row_offset = 0
        tiers = self.get_tiers()
        for child in self.children:
            if filtered and not child.marked_as_filtered:
                continue

            empty = False

            sheet.write(row_offset, 0, "", style_sep)
            for i in range(1, 21):
                sheet.write(row_offset, i, "", style_sep)

            row_offset += 2

            child_index = self.children.index(child)
            if child_index > 0:
                left_context = self.children[child_index - 1]
            else:
                left_context = None
            if left_context:

                sheet.write(row_offset, 0, "Left context", style_context_sep)
                row_offset += 2

                row = 0
                for tier in tiers:
                    sheet.write(row_offset + row, 0, tier, style_tiers)
                    row += 1
                    tier_col[tier] = 1
                print_block(left_context, sheet, row_offset)
                row_offset += len(tiers) + 1
            else:
                sheet.write(row_offset, 0, "No left context", style_context_sep)
                row_offset += 2

            sheet.write(row_offset, 0, "Search result", style_context_sep)
            row_offset += 2

            row = row_offset
            for tier in tiers:
                sheet.write(row, 0, tier, style_tiers)
                row += 1
                tier_col[tier] = 1

            print_block(child, sheet, row_offset)
            row_offset += len(tiers) + 1

            if child_index < len(self.children) - 1:
                right_context = self.children[child_index + 1]
            else:
                right_context = None

            if right_context:

                sheet.write(row_offset, 0, "Right context", style_context_sep)
                row_offset += 2

                row = row_offset
                for tier in tiers:
                    sheet.write(row, 0, tier, style_tiers)
                    row += 1
                    tier_col[tier] = 1
                print_block(right_context, sheet, row_offset)
                row_offset += len(tiers) + 1
            else:
                sheet.write(row_offset, 0, "No right context", style_context_sep)
                row_offset += 1

        if not empty:
            wb.save(output_file)
            return True
        else:
            return False

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

    # marks only those children of root which suit search query
    def filter(self, search):
        # primary check of 'search' dict object, should have fields 'type' and 'value' and 'tiers' optionally
        # 'type' may be 'str', 'substr', 're', 'and', 'or'
        # 'value' for first three types above should have type string, for the last two - list
        # 'tiers' may be specified as a filter list of tier_ids where to search

        '''
        exclude_list = list()

        def process_element(obj):
            string = obj['value']
            for i in range(0,len(string)):
                if string[i] == "\"":
        '''

        if not search:
            return None
        if type(search) != dict or 'type' not in search.keys() or 'value' not in search.keys():
            raise KeyError("Bad type of a search query, should be dict with type and value keys")

        # build search chains from search and compare them with a tree chain
        def search_chain(search, tree_chain):
            if search['type'] not in ('str', 'substr', 're', 'and', 'or'):
                raise KeyError("Bad type of a search query, 'type' key should be 'str', 'substr', 're', 'and' or 'or'")

            if search['type'] == 'and' or search['type'] == 'or':

                if not search['value'] or len(search['value']) == 0:
                    return True
                if type(search['value']) != list:
                    raise KeyError(
                        "Bad value in a search query, if element's type is 'and' or 'or', value should have type list")

                # add all tiers from higher to lower level of the search
                for child_search_value in search['value']:
                    if 'tiers' in search.keys():
                        if 'tiers' not in child_search_value.keys():
                            child_search_value['tiers'] = search['tiers']
                        else:
                            if type(child_search_value['tiers']) != list:
                                raise KeyError(
                                    "Bad value in a search query, element has 'tiers' key, but it's type is not list")
                            for tier in search['tiers']:
                                if tier not in child_search_value['tiers']:
                                    child_search_value['tiers'].append(tier)

                if len(search['value']) == 1:
                    return search_chain(search['value'][0], tree_chain)

                if len(search['value']) > 1:
                    next_search_elem = dict()
                    next_search_elem['value'] = search['value'][1:]
                    if search['type'] == 'or':
                        next_search_elem['type'] = 'or'
                        return search_chain(search['value'][0], tree_chain) or search_chain(next_search_elem,
                                                                                            tree_chain)
                    else:
                        next_search_elem['type'] = 'and'
                        return search_chain(search['value'][0], tree_chain) and search_chain(next_search_elem,
                                                                                             tree_chain)

            else:
                if type(search['value']) != str:
                    raise KeyError(
                        "Bad value in a search query, element's 'type' key value is string-like, but type(element['value']) != str")
                result = False
                for element in tree_chain:

                    if ('tiers' in search.keys() and search['tiers'] and element['tier'] in search['tiers']) or \
                            ('tiers' not in search.keys() or not search['tiers']):

                        if search['type'] == 'substr':
                            if element['data'].find(search['value']) != -1:
                                result = True

                        if search['type'] == 'str':
                            if element['data'] == search['value']:
                                result = True

                        # if search['type'] == 're':

                return result

        # build tree chains and call check function for each of them
        # this variation of function builds chains based on ways from root to leaves
        def tree_chain(node, chain):
            to_append = dict()
            to_append['tier'] = node.tier_id
            to_append['data'] = node.data
            chain.append(to_append)
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
