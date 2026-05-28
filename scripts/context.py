"""Document context: shared state for counters, labels, and bookmarks."""


class DocContext:
    """Tracks heading counters, table/equation/figure numbers, bookmarks, and labels."""

    def __init__(self):
        self.heading_counters = [0, 0, 0, 0]
        self.table_counter = 0
        self.equation_counter = 0
        self.figure_counter = 0
        self._bookmark_id = 0
        self.label_map = {}          # label → (number, bookmark_anchor)
        self.list_counters = [0, 0, 0, 0]
        self.citation_map = {}       # citation_key → number  (for future [@key] support)

    def reset_counters(self):
        """Reset rendering counters for Pass 2 (label_map is preserved)."""
        self.heading_counters = [0, 0, 0, 0]
        self.table_counter = 0
        self.equation_counter = 0
        self.figure_counter = 0
        self._bookmark_id = 0
        self.list_counters = [0, 0, 0, 0]

    def next_bookmark(self, name):
        self._bookmark_id += 1
        return self._bookmark_id

    def register_label(self, label, number, anchor):
        """Map a user label to its number and bookmark anchor."""
        self.label_map[label] = (number, anchor)

    def resolve_label(self, label):
        """Return (number, anchor) for a registered label, or (None, None)."""
        return self.label_map.get(label, (None, None))
