<?php
namespace Netberry;

class XlsxWriter
{
    private static $files = array(
        '_rels/.rels' => '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',

        'xl/_rels/workbook.xml.rels' => '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',

        'xl/worksheets/sheet1.xml' => '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><sheetData>{rows}</sheetData></worksheet>',

        'xl/styles.xml' => '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><charset val="204"/><scheme val="minor"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="0" tint="-0.14999847407452621"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border></border></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" /><xf numFmtId="0" fontId="1" fillId="2" /></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>',

        'xl/workbook.xml' => '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr defaultThemeVersion="124226"/><sheets><sheet name="{SheetName}" sheetId="1" r:id="rId1"/></sheets></workbook>',

        '[Content_Types].xml' => '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    );

    public static function write($filename, $data, array $options = array())
    {
        $options = array_merge(array(
            'sheetName' => 'export',
            'encoding' => 'UTF-8',
        ), $options);

        if (!extension_loaded('zip')) {
            return false;
        }

        $zip = new \ZipArchive();
        if (!$zip->open($filename, \ZIPARCHIVE::CREATE)) {
            return false;
        }

        foreach (self::$files as $srcFilename => $content) {
            $content = '<' . '?xml version="1.0" encoding="UTF-8" standalone="yes"?' . '>' . $content;

            if ($srcFilename == 'xl/workbook.xml') {
                $content = str_replace('{SheetName}', $options['sheetName'], $content);
            } elseif ($srcFilename == 'xl/worksheets/sheet1.xml') {
                $content = str_replace('{rows}', self::generateRows($data, $options), $content);
            }

            $zip->addFromString($srcFilename, $content);
        }

        $zip->close();
        return true;
    }

    private static function generateRows($data, array $options = array())
    {
        if (empty($data) || !is_array($data)) {
            return '';
        }
        $data = array_values($data);

        $keys = array_keys($data[0]);

        $result = '';
        // header
        $result .= '<row>';
        foreach ($keys as $key) {
            $result .= '<c s="1" t="inlineStr"><is><t>' . $key . '</t></is></c>';
        }
        $result .= '</row>';
        // values
        foreach ($data as $row) {
            $result .= '<row>';

            foreach ($keys as $key) {
                $value = (array_key_exists($key, $row)) ? $row[$key] : '';

                if (is_array($value)) {
                    $value = self::flattenArray($value);
                }

                if (preg_match('#^[\d\.]+$#', $value)) {
                    $result .= '<c><v>' . $value . '</v></c>';
                } else {
                    if (defined('ENT_XML1')) {
                        $value = htmlspecialchars($value, ENT_QUOTES | ENT_XML1, $options['encoding']);
                    } else {
                        $value = htmlspecialchars($value);
                    }

                    $result .= '<c t="inlineStr"><is><t>' . $value . '</t></is></c>';
                }
            }

            $result .= '</row>';
        }

        return $result;
    }

    private static function flattenArray($value)
    {
        $result = array();
        foreach ($value as $key => $value2) {
            if (is_array($value2)) {
                $result[] = $key . ' (' . self::flattenArray($value2) . ')';
            } else {
                $result[] = $value2;
            }
        }

        return implode(', ', $result);
    }
}