class Stack:
    def __init__(self):
        self.items = []

    def push(self, item):
        self.items.append(item)

    def pop(self):
        if not self.is_empty():
            return self.items.pop()
        return None

    def peek(self):
        if not self.is_empty():
            return self.items[-1]
        return None

    def is_empty(self):
        return len(self.items) == 0


class Queue:
    def __init__(self):
        self.items = []

    def enqueue(self, item):
        self.items.append(item)

    def dequeue(self):
        if not self.is_empty():
            return self.items.pop(0)
        return None

    def is_empty(self):
        return len(self.items) == 0



class Node:
    def __init__(self, value):
        self.value = value
        self.next = None


class LinkedList:
    def __init__(self):
        self.head = None

    def append(self, value):
        new_node = Node(value)
        if not self.head:
            self.head = new_node
        else:
            current = self.head
            while current.next:
                current = current.next
            current.next = new_node

    def get(self, index):
        current = self.head
        for _ in range(index):
            if not current:
                return None
            current = current.next
        return current

    def update(self, index, value):
        node = self.get(index)
        if node:
            node.value = value


class Graph:
    def __init__(self):
        self.adjacency_list = {}

    def add_edge(self, from_node, to_node):
        if from_node not in self.adjacency_list:
            self.adjacency_list[from_node] = []
        self.adjacency_list[from_node].append(to_node)

    def get_dependents(self, node):
        return self.adjacency_list.get(node, [])

    def remove_node(self, node):
        self.adjacency_list.pop(node, None)
        for dependents in self.adjacency_list.values():
            if node in dependents:
                dependents.remove(node)

    def has_cycle(self):
        visited = set()
        rec_stack = set()

        def dfs(node):
            if node in rec_stack:
                return True
            if node in visited:
                return False
            visited.add(node)
            rec_stack.add(node)
            for neighbor in self.adjacency_list.get(node, []):
                if dfs(neighbor):
                    return True
            rec_stack.remove(node)
            return False

        return any(dfs(node) for node in self.adjacency_list)


