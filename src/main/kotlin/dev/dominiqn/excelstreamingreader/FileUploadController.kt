package dev.dominiqn.excelstreamingreader

import org.slf4j.LoggerFactory
import org.springframework.http.ResponseEntity
import org.springframework.stereotype.Controller
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.ResponseBody
import org.springframework.web.multipart.MultipartFile
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody

@Controller
class FileUploadController(
    private val excelService: ExcelService,
) {
    private val logger = LoggerFactory.getLogger(this::class.java)

    @ResponseBody
    @PostMapping("/upload")
    fun handleFileUpload(
        @RequestParam("file") file: MultipartFile,
        @RequestParam("sheet") sheet: String,
    ): ResponseEntity<StreamingResponseBody> {
        return excelService.handleSheet(file, sheet)
    }
}
