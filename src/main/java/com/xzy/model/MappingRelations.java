package com.xzy.model;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.BiConsumer;

/** 保留关系候选及来源的值映射；键和值按原字符串精确匹配。 */
public final class MappingRelations {
    private static final Match NO_MATCH = new Match(false, Collections.emptyMap());

    private final Map<Pair, LinkedHashSet<Origin>> relations = new LinkedHashMap<>();
    private final Map<String, Candidates> forward = new LinkedHashMap<>();
    private final Map<String, Candidates> reverse = new LinkedHashMap<>();
    private final Map<File, File> sourceFiles = new LinkedHashMap<>();

    /** rowNumber 为 Excel 中从 1 开始的实际行号。 */
    public void add(String source, String target, File file, int rowNumber) {
        Objects.requireNonNull(source, "source");
        Objects.requireNonNull(target, "target");
        Objects.requireNonNull(file, "file");
        if (rowNumber < 1) {
            throw new IllegalArgumentException("来源行号必须从 1 开始");
        }
        File originFile = sourceFiles.get(file);
        if (originFile == null) {
            File normalized = file.toPath().toAbsolutePath().normalize().toFile();
            originFile = sourceFiles.computeIfAbsent(normalized, key -> key);
        }
        addOrigin(source, target, new Origin(originFile, rowNumber));
    }

    private void addOrigin(String source, String target, Origin origin) {
        Pair pair = new Pair(source, target);
        LinkedHashSet<Origin> origins = relations.get(pair);
        if (origins == null) {
            origins = new LinkedHashSet<>();
            relations.put(pair, origins);
            forward.computeIfAbsent(source, key -> new Candidates(false)).values.put(target, origins);
            reverse.computeIfAbsent(target, key -> new Candidates(true)).values.put(source, origins);
        }
        if (origins.add(origin)) {
            forward.get(source).cachedMatch = null;
            reverse.get(target).cachedMatch = null;
        }
    }

    public void addAll(MappingRelations other) {
        Objects.requireNonNull(other, "other");
        if (other == this) {
            return;
        }
        // 分组遍历关系时，保留另一个模型记录的文件首次导入顺序。
        other.sourceFiles.forEach(sourceFiles::putIfAbsent);
        for (Map.Entry<Pair, LinkedHashSet<Origin>> entry : other.relations.entrySet()) {
            Pair pair = entry.getKey();
            for (Origin origin : entry.getValue()) {
                addOrigin(pair.source, pair.target, origin);
            }
        }
    }

    public void clear() {
        relations.clear();
        forward.clear();
        reverse.clear();
        sourceFiles.clear();
    }

    public boolean isEmpty() {
        return relations.isEmpty();
    }

    /** 不同的 source -> target 关系对数量，重复导入或新增来源行不增加此数。 */
    public int size() {
        return relations.size();
    }

    public Set<File> getSourceFiles() {
        return Collections.unmodifiableSet(new LinkedHashSet<>(sourceFiles.keySet()));
    }

    /** 正向存在候选时始终优先正向，即使该方向存在冲突。 */
    public Match resolve(String value) {
        Candidates candidates = forward.get(value);
        if (candidates == null) {
            candidates = reverse.get(value);
        }
        return candidates == null ? NO_MATCH : candidates.match();
    }

    /** 按关系对首次导入顺序遍历；回调中不要修改此模型。 */
    public void forEachPair(BiConsumer<String, String> consumer) {
        Objects.requireNonNull(consumer, "consumer");
        for (Pair pair : relations.keySet()) {
            consumer.accept(pair.source, pair.target);
        }
    }

    /** 先遍历正向冲突，再遍历反向冲突；不为唯一关系构造候选快照。 */
    public void forEachConflict(BiConsumer<String, Match> consumer) {
        Objects.requireNonNull(consumer, "consumer");
        forEachConflict(forward, consumer);
        forEachConflict(reverse, consumer);
    }

    private static void forEachConflict(Map<String, Candidates> index, BiConsumer<String, Match> consumer) {
        for (Map.Entry<String, Candidates> entry : index.entrySet()) {
            if (entry.getValue().values.size() > 1) {
                consumer.accept(entry.getKey(), entry.getValue().match());
            }
        }
    }

    private static final class Candidates {
        private final boolean reverse;
        private final Map<String, LinkedHashSet<Origin>> values = new LinkedHashMap<>();
        private Match cachedMatch;

        private Candidates(boolean reverse) {
            this.reverse = reverse;
        }

        private Match match() {
            if (cachedMatch == null) {
                Map<String, List<Origin>> snapshot = new LinkedHashMap<>();
                for (Map.Entry<String, LinkedHashSet<Origin>> entry : values.entrySet()) {
                    snapshot.put(entry.getKey(), Collections.unmodifiableList(new ArrayList<>(entry.getValue())));
                }
                cachedMatch = new Match(reverse, Collections.unmodifiableMap(snapshot));
            }
            return cachedMatch;
        }
    }

    /** 某次查询的不可修改快照，后续导入或清空模型不会改变它。 */
    public static final class Match {
        private final boolean reverse;
        private final Map<String, List<Origin>> candidates;
        private final String value;

        private Match(boolean reverse, Map<String, List<Origin>> candidates) {
            this.reverse = reverse;
            this.candidates = candidates;
            this.value = candidates.size() == 1 ? candidates.keySet().iterator().next() : null;
        }

        public boolean isAmbiguous() {
            return candidates.size() > 1;
        }

        public String getValue() {
            return value;
        }

        public boolean isReverse() {
            return reverse;
        }

        public Map<String, List<Origin>> getCandidates() {
            return candidates;
        }
    }

    public static final class Origin {
        private final File file;
        private final int rowNumber;

        private Origin(File file, int rowNumber) {
            this.file = file;
            this.rowNumber = rowNumber;
        }

        public File getFile() {
            return file;
        }

        public int getRowNumber() {
            return rowNumber;
        }

        @Override
        public boolean equals(Object other) {
            if (this == other) return true;
            if (!(other instanceof Origin)) return false;
            Origin origin = (Origin) other;
            return rowNumber == origin.rowNumber && file.equals(origin.file);
        }

        @Override
        public int hashCode() {
            return 31 * file.hashCode() + rowNumber;
        }
    }

    private static final class Pair {
        private final String source;
        private final String target;

        private Pair(String source, String target) {
            this.source = source;
            this.target = target;
        }

        @Override
        public boolean equals(Object other) {
            if (this == other) return true;
            if (!(other instanceof Pair)) return false;
            Pair pair = (Pair) other;
            return source.equals(pair.source) && target.equals(pair.target);
        }

        @Override
        public int hashCode() {
            return 31 * source.hashCode() + target.hashCode();
        }
    }
}
