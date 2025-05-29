package dev.dominiqn.excelstreamingreader

import org.slf4j.LoggerFactory
import org.springframework.stereotype.Controller
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestParam
import org.springframework.web.bind.annotation.ResponseBody
import org.springframework.web.multipart.MultipartFile

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
    ) {
        excelService.handleSheet(file, sheet)
    }
}
