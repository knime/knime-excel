/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by KNIME AG, Zurich, Switzerland
 *  Website: http://www.knime.com; Email: contact@knime.com
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License, Version 3, as
 *  published by the Free Software Foundation.
 *
 *  This program is distributed in the hope that it will be useful, but
 *  WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program; if not, see <http://www.gnu.org/licenses>.
 *
 *  Additional permission under GNU GPL version 3 section 7:
 *
 *  KNIME interoperates with ECLIPSE solely via ECLIPSE's plug-in APIs.
 *  Hence, KNIME and ECLIPSE are both independent programs and are not
 *  derived from each other. Should, however, the interpretation of the
 *  GNU GPL Version 3 ("License") under any applicable laws result in
 *  KNIME and ECLIPSE being a combined program, KNIME AG herewith grants
 *  you the additional permission to use and propagate KNIME together with
 *  ECLIPSE with only the license terms in place for ECLIPSE applying to
 *  ECLIPSE and the GNU GPL Version 3 applying for KNIME, provided the
 *  license terms of ECLIPSE themselves allow for the respective use and
 *  propagation of ECLIPSE together with KNIME.
 *
 *  Additional permission relating to nodes for KNIME that extend the Node
 *  Extension (and in particular that are based on subclasses of NodeModel,
 *  NodeDialog, and NodeView) and that only interoperate with KNIME through
 *  standard APIs ("Nodes"):
 *  Nodes are deemed to be separate and independent programs and to not be
 *  covered works.  Notwithstanding anything to the contrary in the
 *  License, the License does not apply to Nodes, you are not required to
 *  license Nodes under the License, and you are granted a license to
 *  prepare and propagate Nodes, in each case even if such Nodes are
 *  propagated with or for interoperation with KNIME.  The owner of a Node
 *  may freely choose the license terms applicable to such Node, including
 *  when such Node is propagated with or for interoperation with KNIME.
 * ---------------------------------------------------------------------
 *
 * History
 *   11 Sep 2023 (Manuel Hotz, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi3.node;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.net.URL;
import java.nio.channels.Channels;
import java.nio.channels.FileChannel;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.UnsupportedFileFormatException;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageRelationshipTypes;
import org.apache.poi.openxml4j.util.ZipFileZipEntrySource;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.junit.jupiter.api.Test;
import org.knime.ext.poi3.node.io.filehandling.excel.reader.read.ExcelUtils;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 *
 * @author Manuel Hotz, KNIME GmbH, Konstanz, Germany
 */
public class ExcelParserPlayground {

    @Test
    void testParsing() throws IOException, OpenXML4JException, SAXException, ParserConfigurationException,
        XMLStreamException, GeneralSecurityException {

        ClassLoader classloader = org.apache.poi.poifs.filesystem.POIFSFileSystem.class.getClassLoader();
        URL res = classloader.getResource("org/apache/poi/poifs/filesystem/POIFSFileSystem.class");
        String path = res.getPath();
        System.out.println("POI Core came from " + path);

        classloader = org.apache.poi.ooxml.POIXMLDocument.class.getClassLoader();
        res = classloader.getResource("org/apache/poi/ooxml/POIXMLDocument.class");
        path = res.getPath();
        System.out.println("POI OOXML came from " + path);

        classloader = org.apache.poi.hslf.usermodel.HSLFSlideShow.class.getClassLoader();
        res = classloader.getResource("org/apache/poi/hslf/usermodel/HSLFSlideShow.class");
        path = res.getPath();
        System.out.println("POI Scratchpad came from " + path);

        classloader = org.apache.xmlbeans.XmlOptions.class.getClassLoader();
        res = classloader.getResource("org/apache/xmlbeans/XmlOptions.class");
        path = res.getPath();
        System.out.println("Xmlbeans XmlOptions came from " + path);

        final var basePath = Path.of("/Users/manuelhotz/Desktop/excelTest/fileFormatTests");
        final String pw = "knime";

        try (final var entries =
            Files.list(basePath).filter(p -> p.getFileName().toString().matches("^data.*"))) {
            final var excelFiles = entries.toList();

            for (final var p : excelFiles) {
                try {
                    System.out.println("File: " + p.getFileName());
                    try (final var channel = Files.newByteChannel(p)) {
                        try (final var parser = XLParser.createParser(channel, pw)) {
                            parser.parse();
                        } catch (EncryptedDocumentException e) {
                            if (pw != null) {
                                throw createPasswordIncorrectException(p, e);
                            } else {
                                throw createPasswordProtectedFileException(p, e);
                            }
                        }
                    }
                } catch (final IOException e) {
                    System.err.println("Exception processing file: " + p.getFileName());
                    e.printStackTrace();
                }
            }
        }
    }

    sealed interface XLParser extends AutoCloseable permits XSSFParser, SSFParser {

        void parse();

        @Override
        void close() throws IOException;

        private static XLParser createParser(final SeekableByteChannel channel, final String password)
            throws EncryptedDocumentException, IOException, GeneralSecurityException, InvalidFormatException {
            final var fileType = ExcelUtils.sniffFileType(channel);
            switch (fileType) {
                case OLE2 -> {
                    return createParserFromOLE2(channel, password);
                }
                case OOXML -> {
                    return new XSSFParser(OPCPackage.open(new ZipFileZipEntrySource(new ZipFile(channel))));
                }
                default -> throw new NotExcelFileException(
                    "File format \"%s\" is not supported for spreadsheet reading.".formatted(fileType));
            }
        }

        static class NotExcelFileException extends UnsupportedFileFormatException {

            protected NotExcelFileException(final String s) {
                super(s);
            }

        }

        static class NotOLE2FileException extends UnsupportedFileFormatException {

            protected NotOLE2FileException(final String s) {
                super(s);
            }

        }

        /**
         * Creates a parser on the given channel containing a valid OLE2 file.
         */
        private static XLParser createParserFromOLE2(final SeekableByteChannel channel, final String password)
                throws IOException, GeneralSecurityException, InvalidFormatException {
            try (final var fs = openFS(channel)) {
                // first try for XML-based files
                if (fs.getRoot().hasEntry(Decryptor.DEFAULT_POIFS_ENTRY)) {
                    // encrypted XLSX/XLSB
                    final var ei = new EncryptionInfo(fs); // this will fail on unencrypted files!
                    final var d = Decryptor.getInstance(ei);
                    if (!d.verifyPassword(password)) {
                        throw new EncryptedDocumentException("Supplied password is invalid");
                    }
                    return new XSSFParser(OPCPackage.open(d.getDataStream(fs))); // this will cache the data in memory
                } else {
                    // then test XLS
                    return new SSFParser(WorkbookFactory.create(fs.getRoot(), password)); // will hold whole sheet in memory
                }
            }
        }

        @SuppressWarnings("resource")
        private static POIFSFileSystem openFS(final SeekableByteChannel channel) throws IOException {
            if (channel instanceof FileChannel fchan) {
                return new POIFSFileSystem(fchan);
            }
            // POIFSFileSystem closes the stream after copying the contents
            return new POIFSFileSystem(new BufferedInputStream(Channels.newInputStream(channel)));
        }
    }

    static final class XSSFParser implements XLParser {

        private OPCPackage m_pkg;

        XSSFParser(final OPCPackage pkg) {
            m_pkg = pkg;
        }

        @Override
        public void close() throws IOException {
            m_pkg.close();
        }

        @Override
        public void parse() {
            var core = m_pkg.getRelationshipsByType(PackageRelationshipTypes.CORE_DOCUMENT);
            if (core.size() == 0) {
                core = m_pkg.getRelationshipsByType(PackageRelationshipTypes.STRICT_CORE_DOCUMENT);
            }
            final var corePart = m_pkg.getPart(core.getRelationship(0));
            final var coreType = corePart.getContentType();

            if (XSSFRelation.XLSB_BINARY_WORKBOOK.getContentType().equals(coreType)) {
                // use binary parser
                System.out.println("\t\t BINARY parser");
            } else {
                // try xml parser
                System.out.println("\t\t XML parser");
            }
        }

    }

    static final class SSFParser implements XLParser {

        private Workbook m_workbook;

        SSFParser(final Workbook workbook) {
            m_workbook = workbook;
        }

        @Override
        public void close() throws IOException {
            m_workbook.close();
        }

        @Override
        public void parse() {
            System.out.println("\t\t Workbook");
        }

    }

    private static void printContents(final OPCPackage opc) throws IOException, OpenXML4JException, XMLStreamException {
        final var reader = new XSSFReader(opc);
        final var fac = XMLInputFactory.newInstance();
        final var sheets = reader.getSheetsData();
        final int max = 1000;
        while (sheets.hasNext()) {
            System.out.println("Procesing new sheet:\n");
            try (final var sheet = sheets.next()) {
                XMLStreamReader streamReader = null;
                try {
                    streamReader = fac.createXMLStreamReader(sheet);
                    for (int i = 0; i < max && streamReader.hasNext(); i++) {
                        System.out.println(mapEvent(streamReader.getEventType()));
                        //                        printEvent(streamReader);
                        streamReader.next();
                    }
                } finally {
                    if (streamReader != null) {
                        streamReader.close();
                    }
                }
            }
        }
    }

    private static void printEvent(final XMLStreamReader xmlr) {

        System.out
            .print("EVENT:[" + xmlr.getLocation().getLineNumber() + ";" + xmlr.getLocation().getColumnNumber() + "] ");

        System.out.print(" [");

        switch (xmlr.getEventType()) {

            case XMLStreamConstants.START_ELEMENT:
                System.out.print("START_ELEMENT ");
                System.out.print("<");
                printName(xmlr);
                printNamespaces(xmlr);
                printAttributes(xmlr);
                System.out.print(">");
                break;

            case XMLStreamConstants.END_ELEMENT:
                System.out.print("</");
                printName(xmlr);
                System.out.print(">");
                break;

            case XMLStreamConstants.SPACE:

            case XMLStreamConstants.CHARACTERS:
                int start = xmlr.getTextStart();
                int length = xmlr.getTextLength();
                System.out.print(new String(xmlr.getTextCharacters(), start, length));
                break;

            case XMLStreamConstants.PROCESSING_INSTRUCTION:
                System.out.println("PROCESSING_INSTRUCTIOn");
                System.out.print("<?");
                if (xmlr.hasText()) {
                    System.out.print(xmlr.getText());
                }
                System.out.print("?>");
                break;

            case XMLStreamConstants.CDATA:
                System.out.print("<![CDATA[");
                start = xmlr.getTextStart();
                length = xmlr.getTextLength();
                System.out.print(new String(xmlr.getTextCharacters(), start, length));
                System.out.print("]]>");
                break;

            case XMLStreamConstants.COMMENT:
                System.out.print("<!--");
                if (xmlr.hasText()) {
                    System.out.print(xmlr.getText());
                }
                System.out.print("-->");
                break;

            case XMLStreamConstants.ENTITY_REFERENCE:
                System.out.print(xmlr.getLocalName() + "=");
                if (xmlr.hasText()) {
                    System.out.print("[" + xmlr.getText() + "]");
                }
                break;

            case XMLStreamConstants.START_DOCUMENT:
                System.out.print("<?xml");
                System.out.print(" version='" + xmlr.getVersion() + "'");
                System.out.print(" encoding='" + xmlr.getCharacterEncodingScheme() + "'");
                if (xmlr.isStandalone()) {
                    System.out.print(" standalone='yes'");
                } else {
                    System.out.print(" standalone='no'");
                }
                System.out.print("?>");
                break;

        }
        System.out.println("]");
    }

    private static void printName(final XMLStreamReader xmlr) {
        if (xmlr.hasName()) {
            String prefix = xmlr.getPrefix();
            String uri = xmlr.getNamespaceURI();
            String localName = xmlr.getLocalName();
            printName(prefix, uri, localName);
        }
    }

    private static void printName(final String prefix, final String uri, final String localName) {
        if (uri != null && !("".equals(uri))) {
            System.out.print("['" + uri + "']:");
        }
        if (prefix != null) {
            System.out.print(prefix + ":");
        }
        if (localName != null) {
            System.out.print(localName);
        }
    }

    private static void printAttributes(final XMLStreamReader xmlr) {
        for (int i = 0; i < xmlr.getAttributeCount(); i++) {
            printAttribute(xmlr, i);
        }
    }

    private static void printAttribute(final XMLStreamReader xmlr, final int index) {
        String prefix = xmlr.getAttributePrefix(index);
        String namespace = xmlr.getAttributeNamespace(index);
        String localName = xmlr.getAttributeLocalName(index);
        String value = xmlr.getAttributeValue(index);
        System.out.print(" ");
        printName(prefix, namespace, localName);
        System.out.print("='" + value + "'");
    }

    private static void printNamespaces(final XMLStreamReader xmlr) {
        for (int i = 0; i < xmlr.getNamespaceCount(); i++) {
            printNamespace(xmlr, i);
        }
    }

    private static void printNamespace(final XMLStreamReader xmlr, final int index) {
        String prefix = xmlr.getNamespacePrefix(index);
        String uri = xmlr.getNamespaceURI(index);
        System.out.print(" ");
        if (prefix == null) {
            System.out.print("xmlns='" + uri + "'");
        } else {
            System.out.print("xmlns:" + prefix + "='" + uri + "'");
        }
    }

    private static void compareParsers(final OPCPackage opc)
        throws IOException, OpenXML4JException, SAXException, ParserConfigurationException, XMLStreamException {
        final var reader = new XSSFReader(opc);
        final var parser = SAXHelper.newXMLReader();
        final var fac = XMLInputFactory.newInstance();
        final var sheets = reader.getSheetsData();
        while (sheets.hasNext()) {
            System.out.println("Procesing new sheet:\n");
            try (final var sheet = sheets.next()) {
                final var streamReader = fac.createXMLEventReader(sheet); // iterator-style
                final var sheetSource = new InputSource(sheet);

                final var handler = new DefaultHandler() {

                    @Override
                    public void characters(final char[] ch, final int start, final int length) throws SAXException {
                        expectNextStaxEvent(XMLStreamConstants.CHARACTERS);
                        super.characters(ch, start, length);
                    }

                    @Override
                    public void startDocument() throws SAXException {
                        expectNextStaxEvent(XMLStreamConstants.START_DOCUMENT);
                        super.startDocument();
                    }

                    @Override
                    public void endDocument() throws SAXException {
                        expectNextStaxEvent(XMLStreamConstants.END_DOCUMENT);
                        super.endDocument();
                    }

                    @Override
                    public void startElement(final String uri, final String localName, final String qName,
                        final Attributes attributes) throws SAXException {
                        expectNextStaxEvent(XMLStreamConstants.START_ELEMENT);
                        super.startElement(uri, localName, qName, attributes);
                    }

                    @Override
                    public void endElement(final String uri, final String localName, final String qName)
                        throws SAXException {
                        expectNextStaxEvent(XMLStreamConstants.END_ELEMENT);
                        super.endElement(uri, localName, qName);
                    }

                    private void expectNextStaxEvent(final int expectedEvent) throws SAXException {
                        try {
                            final var actualEvent = streamReader.nextTag().getEventType();
                            //                            final var actualEvent = streamReader.nextEvent().getEventType();
                            if (actualEvent != expectedEvent) {
                                throw new IllegalStateException(
                                    "Expected %s, got %s%n".formatted(expectedEvent, actualEvent));
                            } else {
                                System.out.println(mapEvent(actualEvent));
                            }

                        } catch (XMLStreamException e) {
                            throw new SAXException(e);
                        }
                    }
                };
                parser.setContentHandler(handler);
                parser.parse(sheetSource);
            }
            System.out.println();
        }
    }

    private static final String mapEvent(final int event) {
        return switch (event) {
            case XMLStreamConstants.ATTRIBUTE -> "ATTRIBUTE";
            case XMLStreamConstants.CDATA -> "CDATA";
            case XMLStreamConstants.CHARACTERS -> "CHARACTERS";
            case XMLStreamConstants.COMMENT -> "COMMENT";
            case XMLStreamConstants.DTD -> "DTD";
            case XMLStreamConstants.END_DOCUMENT -> "END_DOCUMENT";
            case XMLStreamConstants.END_ELEMENT -> "END_ELEMENT";
            case XMLStreamConstants.ENTITY_DECLARATION -> "ENTITY_DECLARATION";
            case XMLStreamConstants.ENTITY_REFERENCE -> "ENTITY_REFERENCE";
            case XMLStreamConstants.NAMESPACE -> "NAMESPACE";
            case XMLStreamConstants.NOTATION_DECLARATION -> "NOTATION_DECLARATION";
            case XMLStreamConstants.PROCESSING_INSTRUCTION -> "PROCESSING_INSTRUCTION";
            case XMLStreamConstants.SPACE -> "SPACE";
            case XMLStreamConstants.START_DOCUMENT -> "START_DOCUMENT";
            case XMLStreamConstants.START_ELEMENT -> "START_ELEMENT";
            default -> throw new IllegalArgumentException("Unexpected value: " + event);
        };
    }

    /**
     * Creates and returns an {@link IOException} with an error message telling the user which file requires a password
     * to be opened.
     *
     * @param e the {@link EncryptedDocumentException} to re-throw, can be {@code null}
     * @return an {@link IOException} with a nice error message
     */
    private static IOException createPasswordProtectedFileException(final Path path,
        final EncryptedDocumentException e) {
        return new IOException(
            String.format("The file '%s' is password protected. Supply a password via the encryption settings.", path),
            e);
    }

    /**
     * Creates and returns an {@link IOException} with an error message telling the user for which file the password is
     * incorrect.
     *
     * @param e the {@link EncryptedDocumentException} to re-throw, can be {@code null}
     * @return an {@link IOException} with a nice error message
     */
    private static IOException createPasswordIncorrectException(final Path path, final EncryptedDocumentException e) {
        return new IOException(String.format("The supplied password is incorrect for file '%s'.", path), e);
    }

}
