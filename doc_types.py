DOC_TEXT = "Text"

IMPORT_FILTER_MAP = {
    "docx": { "FilterName": "Microsoft Word 2007-2013 XML" },
    "doc": { "FilterName": "Microsoft Word 97-2003" }
}

EXPORT_FILTER_MAP = {
    "pdf": {
        DOC_TEXT: {
            "FilterName": "writer_pdf_Export",
            "FilterData": { "IsSkipEmptyPages": True }
        }
    }
}