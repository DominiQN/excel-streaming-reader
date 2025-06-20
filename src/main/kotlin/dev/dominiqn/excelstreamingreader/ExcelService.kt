package dev.dominiqn.excelstreamingreader

import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.util.CellAddress
import org.apache.poi.ss.util.CellReference
import org.apache.poi.util.XMLHelper
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler
import org.apache.poi.xssf.usermodel.XSSFComment
import org.springframework.http.HttpHeaders
import org.springframework.http.MediaType
import org.springframework.http.ResponseEntity
import org.springframework.stereotype.Service
import org.springframework.web.multipart.MultipartFile
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody
import org.xml.sax.InputSource
import java.io.BufferedWriter
import java.io.OutputStreamWriter
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import kotlin.io.path.pathString

@Service
class ExcelService {
    fun handleSheet(file: MultipartFile, sheet: String): ResponseEntity<StreamingResponseBody> {
        // Create a streaming response
        val responseBody = StreamingResponseBody { outputStream ->
            val tempXlsx = Files.createTempFile("excel-streaming-reader-", ".xlsx")

            try {
                // Create a buffered writer that writes directly to the response output stream
                OutputStreamWriter(outputStream, StandardCharsets.UTF_8).buffered().use { output ->
                    // buffer로 8192 bytes 사용
                    // 참고: java.io.InputStream.transferTo
                    file.transferTo(tempXlsx)

                    val opcPackage = OPCPackage.open(tempXlsx.pathString)
                    val xssfReader = XSSFReader(opcPackage)
                    val styles = xssfReader.getStylesTable()

                    xssfReader.setUseReadOnlySharedStringsTable(true)
                    val shareStringsTable = xssfReader.sharedStringsTable
                    val sheetsIterator = xssfReader.sheetIterator

                    while (sheetsIterator.hasNext()) {
                        sheetsIterator.next().use { sheetInputStream ->
                            if (sheetsIterator.sheetName != sheet) {
                                println("sheetName: ${sheetsIterator.sheetName} != $sheet, skip this sheet")
                                return@use
                            }

                            val formatter = DataFormatter(true)
                            val sheetSource = InputSource(sheetInputStream)

                            val parser = XMLHelper.newXMLReader()
                            val contentHandler = XSSFSheetXMLHandler(
                                styles,
                                null,
                                shareStringsTable,
                                SheetHandler(output),
                                formatter,
                                false,
                            )
                            parser.contentHandler = contentHandler
                            parser.parse(sheetSource)
                        }
                    }
                }
            } finally {
                Files.deleteIfExists(tempXlsx)
            }
        }

        // Set up response headers
        val headers = HttpHeaders()
        headers.contentType = MediaType.parseMediaType("text/csv")
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"${file.originalFilename?.replace(".xlsx", ".csv") ?: "export.csv"}\"")

        return ResponseEntity.ok()
            .headers(headers)
            .body(responseBody)
    }

    /**
     * 참고: https://svn.apache.org/repos/asf/poi/trunk/poi-examples/src/main/java/org/apache/poi/examples/xssf/eventusermodel/XLSX2CSV.java
     */
    class SheetHandler(
        private val output: BufferedWriter,
        private val minColumns: Int = -1,
    ) : SheetContentsHandler {
        private var firstCellOfRow = false
        private var currentRow: Int = -1
        private var currentCol: Int = -1

        /**
         * row의 첫 셀 데이터를 읽기 전에 트리거되면서, 해당 row의 row 번호를 알려주는 역할.
         */
        override fun startRow(rowNum: Int) {
            // 만약 차이가 있다면, 부족한 rows만큼 출력
            outputMissingRows(rowNum - currentRow - 1)

            // 현재 row 준비
            firstCellOfRow = true
            currentRow = rowNum
            currentCol = -1
        }

        /**
         * row 마지막 셀 데이터를 읽고 난 후 트리거되면서, 해당 row의 row 번호를 알려주는 역할.
         */
        override fun endRow(rowNum: Int) {
            // 최소 column 길이만큼 보장
            for (i in currentCol..<minColumns) {
                output.append(',')
            }
            output.append('\n')
        }

        override fun cell(
            cellReference: String?,
            formattedValue: String?,
            comment: XSSFComment?,
        ) {
            if (firstCellOfRow) {
                firstCellOfRow = false
            } else {
                output.append(',')
            }

            // cell 주소가 null이라면, 현재 row, col 기준으로 cell 주소 얻기
            val cellRef = cellReference ?: CellAddress(currentRow, currentCol).formatAsString()

            // 빠진 column이 있다면 채워넣기.
            val thisCol = CellReference(cellRef).col.toInt()
            val missedCols = thisCol - currentCol - 1
            for (i in 0..<missedCols) {
                output.append(',')
            }

            // 만약 값이 없다면, 더 이상 추가할 것 없음.
            if (formattedValue == null) {
                return
            }

            currentCol = thisCol

            try {
                formattedValue.toDouble()
                output.append(formattedValue)
            } catch (e: Exception) {
                // 큰따옴표가 있으면, 지우기
                val value = if (formattedValue.startsWith('"') && formattedValue.endsWith('"')) {
                    formattedValue.substring(1, formattedValue.length - 1)
                } else {
                    formattedValue
                }

                output.append('"')
                // CSV 형식을 유효하게 만들기 위해 큰따옴표를 두 개의 큰따옴표로 인코딩
                output.append(value.replace("\"", "\"\""))
                output.append('"')
            }
        }

        private fun outputMissingRows(number: Int) {
            for (i in 0..<number) {
                for (j in 0..<minColumns) {
                    output.append(',')
                }
                output.append('\n')
            }
        }

    }
}
