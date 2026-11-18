package com.xzy.util;

import org.junit.Assume;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;

import static org.junit.Assert.*;

public class SafeExportTest {
    @Rule public TemporaryFolder folder = new TemporaryFolder();

    @Test
    public void failureAfterPartialWritePreservesDestinationAndCleansStagingFile() throws Exception {
        Path old = folder.newFile("existing.xlsx").toPath();
        byte[] original = {7, 8, 9};
        Files.write(old, original);
        Path missing = folder.getRoot().toPath().resolve("new.xlsx");
        for (Path destination : new Path[]{old, missing}) {
            IOException failure = assertThrows(IOException.class, () -> {
                try (SafeExport export = SafeExport.begin(destination.toFile());
                     OutputStream stream = Files.newOutputStream(export.temporaryFile().toPath())) {
                    stream.write(new byte[]{1, 2, 3});
                    throw new IOException("interrupted write");
                }
            });
            assertEquals("interrupted write", failure.getMessage());
            assertArrayEquals(original, Files.readAllBytes(old));
            assertFalse(Files.exists(missing));
            assertEquals(1, folder.getRoot().list().length);
        }
    }

    @Test
    public void rejectsHardLinkToInput() throws Exception {
        Path source = folder.newFile("source.xlsx").toPath();
        Files.write(source, new byte[]{42});
        Path alias = source.resolveSibling("alias.xlsx");
        try {
            Files.createLink(alias, source);
        } catch (UnsupportedOperationException | IOException e) {
            Assume.assumeNoException("此文件系统不能创建硬链接", e);
        }
        assertThrows(IOException.class, () -> {
            try (SafeExport export = SafeExport.begin(alias.toFile(), source.toFile())) {
                fail("应在创建临时文件之前拒绝输入的硬链接");
            }
        });
        assertTrue(Files.isSameFile(source, alias));
        assertArrayEquals(new byte[]{42}, Files.readAllBytes(source));
    }

    @Test
    public void failedReplacementOnWindowsLeavesOldFileAndRemovesStagingFile() throws Exception {
        Assume.assumeTrue(System.getProperty("os.name").toLowerCase(Locale.ROOT).startsWith("windows"));
        Path destination = folder.newFile("locked.xlsx").toPath();
        byte[] original = {4, 5, 6};
        Files.write(destination, original);
        Path staged;
        // FileInputStream 在 Windows 持有不共享删除权限的句柄，实际替换会失败。
        try (FileInputStream lock = new FileInputStream(destination.toFile());
             SafeExport export = SafeExport.begin(destination.toFile())) {
            staged = export.temporaryFile().toPath();
            Files.write(staged, new byte[]{1, 2});
            assertThrows(IOException.class, export::commit);
        }
        assertArrayEquals(original, Files.readAllBytes(destination));
        assertFalse(Files.exists(staged));
    }
}
