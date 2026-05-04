# Docx MCP Service

A Docx document processing service based on the FastMCP library, supporting the creation, editing, and management of Word documents using AI assistants in Cursor.

## Features

- **Complete Document Operations**: Support for creating, opening, saving documents, as well as adding, editing, and deleting content
- **Formatting**: Support for setting fonts, colors, sizes, alignment, and other formatting options
- **Table Processing**: Support for creating, editing, merging, and splitting table cells
- **Image Insertion**: Support for inserting images and setting their sizes
- **Layout Control**: Support for setting page margins, adding page breaks, and other layout elements
- **Query Functions**: Support for retrieving document information, paragraph content, and table data
- **Convenient Editing**: Support for find and replace functionality
- **Section Editing**: Support for replacing content in specific sections while preserving original formatting and styles

## Installation Dependencies

Ensure Python 3.10+ is installed, then install the following dependencies:

```bash
pip3 install python-docx mcp
```

## Usage

### Server Flags

The server accepts two optional flags to restrict its behavior:

| Flag | Description |
|---|---|
| `--readonly` | Disables all write operations. Open, search, and read tools still work. Useful for inspecting documents safely. |
| `--cwd-only` | Restricts all file access to the current working directory. Paths outside it are rejected. |

Flags can be combined:

```bash
# Read-only mode
python3 server.py --readonly

# Restrict to current directory only
python3 server.py --cwd-only

# Both at once
python3 server.py --readonly --cwd-only
```

---

### Using as an MCP Service in Cursor

1. Open Cursor and go to Settings
2. Find the `Features > MCP Servers` section
3. Click `Add new MCP server`
4. Fill in the following information:
   - Name: MCP_DOCX
   - Type: Command
   - Command: `python3 /path/to/MCP-Doc/server.py` (add `--readonly` or `--cwd-only` flags as needed)
5. Click `Add` to add the service

After adding, you can use natural language to operate Word documents in Cursor's AI assistant, for example:

- "Create a new Word document and save it to the desktop"
- "Add a level 3 heading"
- "Insert a 3x4 table and fill it with data"
- "Set the second paragraph to bold and center-aligned"

### Using as an MCP Service in OpenCode

[OpenCode](https://github.com/opencode-ai/opencode) supports MCP servers via its configuration file.

1. Locate or create the OpenCode configuration file at `~/.config/opencode/config.json`
2. Add the MCP server entry under the `mcpServers` key:

```json
{
  "mcp": {
    "MCP_DOCX": {
      "type": "local",
      "command": ["python", "/path/to/MCP-Doc/server.py", "--readonly"]
    }
  }
}
```

To start with flags, add them to the command array:

```json
{
  "mcpServers": {
    "MCP_DOCX": {
      "type": "local",
      "command": ["python3", "/path/to/MCP-Doc/server.py", "--readonly", "--cwd-only"]
    }
  }
}
```

Replace `/path/to/MCP-Doc/server.py` with the actual absolute path to `server.py`.

3. Restart OpenCode — the MCP_DOCX tools will be available automatically in any session.

## Supported Operations

The service supports the following operations:

- **Document Management**: `create_document`, `open_document`, `save_document`
- **Content Addition**: `add_paragraph`, `add_heading`, `add_table`, `add_picture`
- **Content Editing**: `edit_paragraph`, `delete_paragraph`, `delete_text`
- **Table Operations**: `add_table_row`, `delete_table_row`, `edit_table_cell`, `merge_table_cells`, `split_table`
- **Layout Control**: `add_page_break`, `set_page_margins`
- **Query Functions**: `get_document_info`, `get_headings`, `get_full_text`, `get_paragraph`, `get_paragraphs`, `get_table_content`, `get_section_content`, `get_document_outline`, `search_text`
- **File Operations**: `create_document`, `open_document`, `save_document`, `save_as_document`, `create_document_copy`
- **Section Editing**: `replace_section`, `edit_section_by_keyword`
- **Other Functions**: `find_and_replace`, `search_and_replace` (with preview functionality)

## How It Works

1. The service uses the Python-docx library to process Word documents
2. It implements the MCP protocol through the FastMCP library to communicate with AI assistants
3. It processes requests and returns formatted responses
4. It supports complete error handling and status reporting

## Typography Capabilities

The service has good typography understanding capabilities:

- **Text Hierarchy**: Support for heading levels (1-9) and paragraph organization
- **Page Layout**: Support for page margin settings
- **Visual Elements**: Support for font styles (bold, italic, underline, color) and alignment
- **Table Layout**: Support for creating tables, merging cells, splitting tables, and setting table formats
- **Pagination Control**: Support for adding page breaks

## Development Notes

- `server.py` - Core implementation of the MCP service using the FastMCP library

## Troubleshooting

If you encounter problems in Cursor, try the following steps:

1. Ensure Python 3.10+ is correctly installed
2. Ensure the python-docx and mcp libraries are correctly installed
3. Check if the server path is correct
4. Restart the Cursor application

## Notes

- Ensure the python-docx and mcp libraries are correctly installed
- Using absolute paths can avoid path parsing issues

## License

MIT License
