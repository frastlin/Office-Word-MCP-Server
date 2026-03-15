"""
MCP tool implementations for the Word Document Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""

# Document tools
from word_document_server.tools.document_tools import (
    create_document, get_document_info, get_document_text, 
    get_document_outline, list_available_documents, 
    copy_document, merge_documents
)

# Content tools
from word_document_server.tools.content_tools import (
    add_heading, add_paragraph, add_table, add_picture,
    add_page_break, add_table_of_contents, delete_paragraph,
    search_and_replace
)

# Format tools
from word_document_server.tools.format_tools import (
    format_text, create_custom_style, format_table
)

# Protection tools
from word_document_server.tools.protection_tools import (
    protect_document, add_restricted_editing,
    add_digital_signature, verify_document
)

# Footnote tools
from word_document_server.tools.footnote_tools import (
    add_footnote_to_document, add_endnote_to_document,
    convert_footnotes_to_endnotes_in_document, customize_footnote_style
)

# Comment tools
from word_document_server.tools.comment_tools import (
    get_all_comments, get_comments_by_author, get_comments_for_paragraph
)

# Track changes tools
from word_document_server.tools.track_changes_tools import (
    replace_with_track_changes, delete_with_track_changes,
    insert_after_with_track_changes, insert_before_with_track_changes,
    list_revisions, accept_revision, reject_revision,
    accept_all_revisions, reject_all_revisions,
    get_visible_text, count_tracked_matches,
)

# Comment management tools
from word_document_server.tools.comment_management_tools import (
    add_comment, reply_to_comment, resolve_comment, delete_comment,
)
