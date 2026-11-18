package com.xzy.util;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

/** 将一次导出暂存在目标目录，写入成功后一次性发布。 */
public final class SafeExport implements Closeable {
    private final Path destination;
    private final Path temporary;
    private boolean committed;

    private SafeExport(Path destination, Path temporary) {
        this.destination = destination;
        this.temporary = temporary;
    }

    public static SafeExport begin(File output, File... inputs) throws IOException {
        Path requested = output.toPath().toAbsolutePath();
        if (requested.getFileName() == null || Files.isDirectory(requested)) {
            throw new IOException("请选择输出文件，不能选择目录: " + output);
        }
        Files.createDirectories(requested.getParent());
        // 解析目录别名；若结果文件本身是链接，保留链接并写入它指向的文件。
        Path destination = Files.isSymbolicLink(requested) ? requested.toRealPath()
                : requested.getParent().toRealPath().resolve(requested.getFileName());
        boolean exists = Files.exists(destination);
        for (File input : inputs) {
            if (input == null) {
                continue;
            }
            Path source = input.toPath();
            if (destination.equals(source.toAbsolutePath().normalize())
                    || (exists && Files.exists(source) && Files.isSameFile(destination, source))) {
                throw new IOException("输出文件不能与输入文件相同: " + input);
            }
        }
        if (exists && (!Files.isRegularFile(destination) || !Files.isWritable(destination))) {
            throw new IOException("无法替换输出文件: " + destination);
        }
        return new SafeExport(destination, Files.createTempFile(destination.getParent(), ".excel-export-", ".tmp"));
    }

    public File temporaryFile() {
        return temporary.toFile();
    }

    /** 只有处理逻辑以及所有写入流的关闭均已成功，才可调用。 */
    public void commit() throws IOException {
        try {
            Files.move(temporary, destination, StandardCopyOption.ATOMIC_MOVE, StandardCopyOption.REPLACE_EXISTING);
            committed = true;
        } catch (AtomicMoveNotSupportedException e) {
            throw new IOException("输出目录不支持安全替换，请选择其他目录: " + destination, e);
        }
    }

    @Override
    public void close() throws IOException {
        if (!committed) {
            Files.deleteIfExists(temporary);
        }
    }
}
