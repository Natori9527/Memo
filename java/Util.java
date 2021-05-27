
package com.example.demo.getFile.service;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import lombok.extern.slf4j.Slf4j;

/**
 * ファイル取得のService.
 *
 */
@Slf4j
@Service
public class GetFileService {

    /**
     * プレビュー表示
     *
     * @param filePath
     *            
     * @return HttpEntity
     */
    public ResponseEntity<byte[]> showFile(String filePath, String mediaType)
            throws FileNotFoundException, IOException {

        byte[] fileContents = null;

        Path path = Paths.get(filePath);
        fileContents = Files.readAllBytes(path);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(mediaType));
        headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");

        return new ResponseEntity<>(fileContents, headers, HttpStatus.OK);
    }
}
