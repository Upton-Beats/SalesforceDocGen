/**
 * DocGenExcelService
 *
 * Generates Excel (.xlsx) documents from templates by replacing {FieldName}
 * merge tags with record data. Preserves all formatting, images, headers,
 * footers, and styles through a ZIP round-trip.
 *
 * The .xlsx format is a ZIP archive containing XML parts. This class uses
 * Compression.ZipReader to unpack the template, performs string replacements
 * on the worksheet XML and shared strings table, then repacks with
 * Compression.ZipWriter.
 *
 * Tag syntax matches DocGen conventions:
 *   {FieldName}              - Simple field replacement
 *   {Relationship.FieldName} - Parent relationship traversal
 *   {FieldName:dateFormat}   - Date formatting
 *   {#RelationshipName}...{/RelationshipName} - Row loops for child records
 *
 * Images in headers/footers are preserved because non-XML ZIP entries
 * (e.g. xl/media/image1.png) are copied byte-for-byte.
 */
public with sharing class DocGenExcelService {

    private static final Integer MAX_TEMPLATE_SIZE = 6000000; // ~6 MB

    /**
     * Processes an Excel template with record data and returns the merged .xlsx Blob.
     *
     * @param templateBlob  The .xlsx template file as a Blob
     * @param recordDataMap Record data as a flat or nested Map (from DocGenDataRetriever)
     * @return Modified .xlsx Blob with all merge tags replaced
     */
    public static Blob processExcelTemplate(Blob templateBlob, Map<String, Object> recordDataMap) {
        if (templateBlob == null) {
            throw new DocGenException('Excel template file is missing.');
        }
        if (templateBlob.size() > MAX_TEMPLATE_SIZE) {
            throw new DocGenException(
                'Excel template exceeds the ' + (MAX_TEMPLATE_SIZE / 1000000) +
                ' MB size limit. Reduce the file size and try again.'
            );
        }
        if (recordDataMap == null) {
            recordDataMap = new Map<String, Object>();
        }

        Compression.ZipReader reader = new Compression.ZipReader(templateBlob);
        Compression.ZipWriter writer = new Compression.ZipWriter();

        for (Compression.ZipEntry entry : reader.getEntries()) {
            String entryName = entry.getName();
            Blob entryBlob = reader.extract(entryName);

            if (isReplaceableEntry(entryName)) {
                String xml = entryBlob.toString();
                xml = processExcelXml(xml, recordDataMap, entryName);
                writer.addEntry(entryName, Blob.valueOf(xml));
            } else {
                // All other entries (styles, images, rels, etc.) - copy unchanged
                writer.addEntry(entryName, entryBlob);
            }
        }

        return writer.getArchive();
    }

    /**
     * Returns true for ZIP entries that may contain merge tags.
     * We process worksheet XML files and the shared strings table.
     */
    private static Boolean isReplaceableEntry(String entryName) {
        if (entryName == 'xl/sharedStrings.xml') {
            return true;
        }
        return entryName.startsWith('xl/worksheets/sheet') && entryName.endsWith('.xml');
    }

    /**
     * Processes Excel XML content, replacing merge tags with record data.
     * Handles simple tags, date-formatted tags, and row loops for child records.
     */
    private static String processExcelXml(String xml, Map<String, Object> data, String entryName) {
        // First pass: handle row loops {#RelName}...{/RelName} in worksheet XML
        if (entryName.startsWith('xl/worksheets/')) {
            xml = processRowLoops(xml, data);
        }

        // Second pass: replace simple {tag} and {tag:format} merge tags
        xml = replaceSimpleTags(xml, data);

        return xml;
    }

    /**
     * Finds {#RelationshipName}...{/RelationshipName} loop blocks in worksheet XML.
     * For each block, identifies the enclosing <row> element(s) and repeats them
     * for each child record.
     */
    private static String processRowLoops(String xml, Map<String, Object> data) {
        Integer cursor = 0;
        String result = '';

        while (cursor < xml.length()) {
            // Find next loop start tag {#...}
            Integer loopTagStart = findNextLoopTag(xml, cursor);
            if (loopTagStart == -1) {
                result += xml.substring(cursor);
                break;
            }

            // Append everything before the loop tag
            Integer tagOpenBrace = xml.lastIndexOf('{', loopTagStart);

            // Extract the loop key
            Integer tagClose = xml.indexOf('}', loopTagStart);
            if (tagClose == -1) {
                result += xml.substring(cursor);
                break;
            }
            String loopKey = xml.substring(loopTagStart + 1, tagClose).trim(); // e.g. "#ContactList"
            String key = loopKey.substring(1); // Remove '#'

            // Find matching end tag {/key}
            String endTag = '{/' + key + '}';
            Integer endTagStart = xml.indexOf(endTag, tagClose);
            if (endTagStart == -1) {
                // No matching end tag - skip this tag
                result += xml.substring(cursor, tagClose + 1);
                cursor = tagClose + 1;
                continue;
            }

            // Find the enclosing <row> element for the loop start
            Integer rowStart = xml.lastIndexOf('<row ', tagOpenBrace);
            if (rowStart == -1) {
                rowStart = xml.lastIndexOf('<row>', tagOpenBrace);
            }

            // Find the </row> after the end tag
            Integer rowEnd = xml.indexOf('</row>', endTagStart);
            if (rowEnd == -1) {
                // No enclosing row - treat as simple replacement
                result += xml.substring(cursor, tagClose + 1);
                cursor = tagClose + 1;
                continue;
            }
            rowEnd += 6; // Include </row>

            // If the row start is before our current cursor, we can't backtrack
            if (rowStart < cursor) {
                result += xml.substring(cursor, tagClose + 1);
                cursor = tagClose + 1;
                continue;
            }

            // Append everything before the row
            result += xml.substring(cursor, rowStart);

            // Extract the row template (without loop tags)
            String rowTemplate = xml.substring(rowStart, rowEnd);
            String startTagFull = '{' + loopKey + '}';
            String endTagFull = endTag;
            rowTemplate = rowTemplate.replace(startTagFull, '');
            rowTemplate = rowTemplate.replace(endTagFull, '');

            // Resolve child records
            Object val = resolveValue(data, key);
            List<Object> records = extractRecordList(val);

            if (records != null && !records.isEmpty()) {
                // Need to renumber rows for valid Excel XML
                Integer baseRowNum = extractRowNumber(rowTemplate);
                Integer rowIdx = 0;
                for (Object item : records) {
                    if (item instanceof Map<String, Object>) {
                        Map<String, Object> childData = (Map<String, Object>) item;
                        String expandedRow = replaceSimpleTags(rowTemplate, childData);
                        // Update row number and cell references
                        if (baseRowNum > 0) {
                            expandedRow = updateRowNumber(expandedRow, baseRowNum, baseRowNum + rowIdx);
                        }
                        result += expandedRow;
                        rowIdx++;
                    }
                }
            }
            // If no records, the row is omitted (empty loop)

            cursor = rowEnd;
        }

        return result;
    }

    /**
     * Finds the next {#...} loop start tag in the XML starting from the given position.
     * Returns the position of the '#' character, or -1 if not found.
     */
    private static Integer findNextLoopTag(String xml, Integer fromPos) {
        Integer pos = fromPos;
        while (pos < xml.length()) {
            Integer braceOpen = xml.indexOf('{', pos);
            if (braceOpen == -1) return -1;

            // Check if next non-space char after { is #
            Integer checkPos = braceOpen + 1;
            while (checkPos < xml.length() && xml.substring(checkPos, checkPos + 1) == ' ') {
                checkPos++;
            }
            if (checkPos < xml.length() && xml.substring(checkPos, checkPos + 1) == '#') {
                return checkPos;
            }
            pos = braceOpen + 1;
        }
        return -1;
    }

    /**
     * Extracts the row number from a <row r="N"> attribute in SpreadsheetML.
     */
    private static Integer extractRowNumber(String rowXml) {
        Integer rIdx = rowXml.indexOf('<row ');
        if (rIdx == -1) rIdx = rowXml.indexOf('<row>');
        if (rIdx == -1) return -1;

        Integer rAttr = rowXml.indexOf(' r="', rIdx);
        if (rAttr == -1) return -1;
        rAttr += 4;
        Integer rEnd = rowXml.indexOf('"', rAttr);
        if (rEnd == -1) return -1;

        String numStr = rowXml.substring(rAttr, rEnd);
        return numStr.isNumeric() ? Integer.valueOf(numStr) : -1;
    }

    /**
     * Updates row numbers and cell references in a row XML fragment.
     * Changes row r="oldNum" to r="newNum" and cell references like A5 to A7.
     */
    private static String updateRowNumber(String rowXml, Integer oldNum, Integer newNum) {
        if (oldNum == newNum) return rowXml;
        String oldStr = String.valueOf(oldNum);
        String newStr = String.valueOf(newNum);

        // Update <row r="N"> attribute
        rowXml = rowXml.replace(' r="' + oldStr + '"', ' r="' + newStr + '"');

        // Update cell references like r="A5" to r="A7"
        // SpreadsheetML cell references are in the format r="[A-Z]+N"
        Pattern cellRefPattern = Pattern.compile(' r="([A-Z]+)' + oldStr + '"');
        Matcher m = cellRefPattern.matcher(rowXml);
        String updated = '';
        Integer lastEnd = 0;
        while (m.find()) {
            updated += rowXml.substring(lastEnd, m.start());
            updated += ' r="' + m.group(1) + newStr + '"';
            lastEnd = m.end();
        }
        if (lastEnd > 0) {
            updated += rowXml.substring(lastEnd);
            return updated;
        }
        return rowXml;
    }

    /**
     * Extracts a list of records from a resolved value.
     * Handles both List<Object> and Map with 'records' key.
     */
    private static List<Object> extractRecordList(Object val) {
        if (val instanceof List<Object>) {
            return (List<Object>) val;
        }
        if (val instanceof Map<String, Object>) {
            Map<String, Object> mapVal = (Map<String, Object>) val;
            if (mapVal.containsKey('records') && mapVal.get('records') instanceof List<Object>) {
                return (List<Object>) mapVal.get('records');
            }
        }
        return null;
    }

    /**
     * Replaces simple {tag} and {tag:format} merge tags in XML content.
     * Uses the same tag syntax as DocGenService for consistency.
     */
    private static String replaceSimpleTags(String xml, Map<String, Object> data) {
        String output = '';
        Integer len = xml.length();
        Integer cursor = 0;

        while (cursor < len) {
            Integer nextTag = xml.indexOf('{', cursor);
            if (nextTag == -1) {
                output += xml.substring(cursor);
                break;
            }
            output += xml.substring(cursor, nextTag);
            cursor = nextTag;

            Integer tagClose = xml.indexOf('}', cursor);
            if (tagClose == -1) {
                output += xml.substring(cursor);
                break;
            }

            String rawTag = xml.substring(cursor + 1, tagClose);
            String tagContent = rawTag.trim();

            // Skip loop tags (handled separately) and closing tags
            if (tagContent.startsWith('#') || tagContent.startsWith('/')) {
                output += '{' + rawTag + '}';
                cursor = tagClose + 1;
                continue;
            }

            // Parse optional format suffix: {FieldName:format}
            String fieldName = tagContent;
            String dateFormat = null;
            if (tagContent.contains(':')) {
                Integer colonIdx = tagContent.indexOf(':');
                fieldName = tagContent.substring(0, colonIdx).trim();
                dateFormat = tagContent.substring(colonIdx + 1).trim();
            }

            Object val = resolveValue(data, fieldName);
            if (val != null) {
                String strVal;
                if (dateFormat != null && val instanceof DateTime) {
                    strVal = ((DateTime) val).format(dateFormat);
                } else if (dateFormat != null && val instanceof Date) {
                    DateTime dt = DateTime.newInstance((Date) val, Time.newInstance(0, 0, 0, 0));
                    strVal = dt.format(dateFormat);
                } else if (val instanceof Date) {
                    Date d = (Date) val;
                    String mm = (d.month() < 10 ? '0' : '') + d.month();
                    String dd = (d.day() < 10 ? '0' : '') + d.day();
                    strVal = mm + '/' + dd + '/' + d.year();
                } else if (val instanceof DateTime) {
                    strVal = ((DateTime) val).format('MM/dd/yyyy HH:mm');
                } else {
                    strVal = String.valueOf(val);
                }
                output += escapeXml(strVal);
            }
            // If val is null, the tag is removed (replaced with empty string)
            cursor = tagClose + 1;
        }

        return output;
    }

    /**
     * Resolves a dotted field path (e.g. "Account.Name") against the data map.
     * Uses case-insensitive matching as a fallback.
     */
    private static Object resolveValue(Map<String, Object> data, String path) {
        if (data == null) return null;
        List<String> parts = path.split('\\.');
        Object current = data;

        for (String part : parts) {
            if (current instanceof Map<String, Object>) {
                Map<String, Object> currentMap = (Map<String, Object>) current;
                if (currentMap.containsKey(part)) {
                    current = currentMap.get(part);
                } else {
                    // Case-insensitive fallback
                    Boolean found = false;
                    for (String key : currentMap.keySet()) {
                        if (key.equalsIgnoreCase(part)) {
                            current = currentMap.get(key);
                            found = true;
                            break;
                        }
                    }
                    if (!found) return null;
                }
            } else if (current instanceof SObject) {
                SObject sobj = (SObject) current;
                try {
                    current = sobj.get(part);
                } catch (Exception e) {
                    return null;
                }
            } else {
                return null;
            }
            if (current == null) return null;
        }
        return current;
    }

    /**
     * Escapes XML special characters. The & replacement must come first
     * to avoid double-encoding.
     */
    private static String escapeXml(String value) {
        if (String.isBlank(value)) return value;
        return value
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;');
    }
}
