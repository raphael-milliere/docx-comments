"""Exception types for docx-comments.

Subclassing the previously raised builtins keeps backward compatibility:
existing ``except ValueError`` / ``except KeyError`` consumers keep working,
while new code can catch the precise types.
"""


class CommentNotFoundError(ValueError, LookupError):
    """Raised when a comment id does not match any comment in the document."""


class PersonNotFoundError(KeyError):
    """Raised when an author has no entry in word/people.xml."""
