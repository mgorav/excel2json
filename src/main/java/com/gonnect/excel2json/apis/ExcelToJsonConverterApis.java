package com.gonnect.excel2json.apis;

import com.gonnect.excel2json.config.ExcelToJsonConverterConfig;
import com.gonnect.excel2json.services.ExcelToJsonConversionService;
import io.swagger.annotations.Api;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@Api("Excel To Json Converter")
public class ExcelToJsonConverterApis {

    @RequestMapping(value = "/excel2json/converter" , method = RequestMethod.POST)
    public String toJson(@RequestBody ExcelToJsonConverterConfig config) throws IOException, InvalidFormatException {

        return ExcelToJsonConversionService.convert(config).toJson();
    }
}
