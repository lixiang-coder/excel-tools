package com.xzy.model;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.junit.Assert.*;

public class MappingRelationsTest {
    @Rule
    public TemporaryFolder temporary = new TemporaryFolder();

    @Test
    public void repeatedPairsDoNotCreateConflictsAndRepeatedOriginsAreDeduplicated() {
        MappingRelations mappings = new MappingRelations();
        File first = file("first.xlsx");
        File second = file("second.xlsx");
        mappings.add("A", "X", first, 2);
        mappings.add("A", "X", first, 2);
        mappings.add("A", "X", first, 3);
        mappings.add("A", "X", second, 2);

        assertEquals(1, mappings.size());
        MappingRelations.Match match = mappings.resolve("A");
        assertFalse(match.isAmbiguous());
        assertFalse(match.isReverse());
        assertEquals("X", match.getValue());
        assertEquals(Arrays.asList("first.xlsx:2", "first.xlsx:3", "second.xlsx:2"),
                originLabels(match.getCandidates().get("X")));
        assertEquals(Arrays.asList(first, second), new ArrayList<>(mappings.getSourceFiles()));
    }

    @Test
    public void forwardConflictsKeepAllCandidatesAndTheirOwnOriginsInOrder() {
        MappingRelations mappings = new MappingRelations();
        mappings.add("A", "second", file("first.xlsx"), 2);
        mappings.add("A", "first", file("second.xlsx"), 8);
        mappings.add("A", "second", file("third.xlsx"), 4);

        MappingRelations.Match match = mappings.resolve("A");

        assertTrue(match.isAmbiguous());
        assertNull(match.getValue());
        assertFalse(match.isReverse());
        assertEquals(Arrays.asList("second", "first"), new ArrayList<>(match.getCandidates().keySet()));
        assertEquals(Arrays.asList("first.xlsx:2", "third.xlsx:4"), originLabels(match.getCandidates().get("second")));
        assertEquals(Collections.singletonList("second.xlsx:8"), originLabels(match.getCandidates().get("first")));
        assertEquals(2, mappings.size());
    }

    @Test
    public void reverseLookupFindsUniqueValuesAndReportsAllConflictingSources() {
        MappingRelations mappings = new MappingRelations();
        mappings.add("A", "X", file("first.xlsx"), 2);
        MappingRelations.Match unique = mappings.resolve("X");
        assertTrue(unique.isReverse());
        assertFalse(unique.isAmbiguous());
        assertEquals("A", unique.getValue());

        mappings.add("B", "X", file("second.xlsx"), 6);
        MappingRelations.Match conflict = mappings.resolve("X");

        assertTrue(conflict.isReverse());
        assertTrue(conflict.isAmbiguous());
        assertNull(conflict.getValue());
        assertEquals(Arrays.asList("A", "B"), new ArrayList<>(conflict.getCandidates().keySet()));
        assertEquals(Collections.singletonList("first.xlsx:2"), originLabels(conflict.getCandidates().get("A")));
        assertEquals(Collections.singletonList("second.xlsx:6"), originLabels(conflict.getCandidates().get("B")));
    }

    @Test
    public void aUniqueForwardMatchTakesPriorityOverReverseConflicts() {
        MappingRelations mappings = new MappingRelations();
        mappings.add("A", "X", file("source.xlsx"), 2);
        mappings.add("B", "A", file("source.xlsx"), 3);
        mappings.add("C", "A", file("source.xlsx"), 4);

        MappingRelations.Match match = mappings.resolve("A");

        assertEquals("X", match.getValue());
        assertFalse(match.isReverse());
        assertFalse(match.isAmbiguous());
    }

    @Test
    public void anAmbiguousForwardMatchDoesNotFallBackToAUniqueReverseMatch() {
        MappingRelations mappings = new MappingRelations();
        mappings.add("A", "X", file("source.xlsx"), 2);
        mappings.add("A", "Y", file("source.xlsx"), 3);
        mappings.add("B", "A", file("source.xlsx"), 4);

        MappingRelations.Match match = mappings.resolve("A");

        assertTrue(match.isAmbiguous());
        assertFalse(match.isReverse());
        assertNull(match.getValue());
        assertEquals(Arrays.asList("X", "Y"), new ArrayList<>(match.getCandidates().keySet()));
    }

    @Test
    public void callbacksVisitPairsOnceAndIncludeBothDirectionsOfConflicts() {
        MappingRelations mappings = new MappingRelations();
        mappings.add("A", "X", file("source.xlsx"), 2);
        mappings.add("B", "X", file("source.xlsx"), 3);
        mappings.add("A", "Y", file("source.xlsx"), 4);
        mappings.add("X", "P", file("source.xlsx"), 5);
        mappings.add("X", "Q", file("source.xlsx"), 6);
        mappings.add("A", "X", file("source.xlsx"), 7);
        List<String> pairs = new ArrayList<>();
        List<String> conflicts = new ArrayList<>();

        mappings.forEachPair((source, target) -> pairs.add(source + "->" + target));
        mappings.forEachConflict((key, match) -> {
            assertTrue(match.isAmbiguous());
            assertNull(match.getValue());
            conflicts.add((match.isReverse() ? "reverse:" : "forward:") + key);
        });

        assertEquals(Arrays.asList("A->X", "B->X", "A->Y", "X->P", "X->Q"), pairs);
        assertEquals(Arrays.asList("forward:A", "forward:X", "reverse:X"), conflicts);
    }

    @Test
    public void addAllPreservesPairCandidateOriginAndSourceFileOrder() {
        MappingRelations incoming = new MappingRelations();
        incoming.add("A", "Z", file("first.xlsx"), 2);
        incoming.add("B", "Y", file("second.xlsx"), 3);
        incoming.add("A", "Y", file("third.xlsx"), 4);
        incoming.add("A", "Z", file("fourth.xlsx"), 5);
        MappingRelations mappings = new MappingRelations();
        mappings.add("C", "Y", file("existing.xlsx"), 2);

        mappings.addAll(incoming);
        mappings.addAll(incoming);
        mappings.addAll(mappings);

        assertEquals(4, mappings.size());
        assertEquals(Arrays.asList("C", "B", "A"), new ArrayList<>(mappings.resolve("Y").getCandidates().keySet()));
        assertEquals(Arrays.asList("first.xlsx:2", "fourth.xlsx:5"),
                originLabels(mappings.resolve("A").getCandidates().get("Z")));
        assertEquals(Arrays.asList(file("existing.xlsx"), file("first.xlsx"), file("second.xlsx"),
                file("third.xlsx"), file("fourth.xlsx")), new ArrayList<>(mappings.getSourceFiles()));
        List<String> pairs = new ArrayList<>();
        mappings.forEachPair((source, target) -> pairs.add(source + "->" + target));
        assertEquals(Arrays.asList("C->Y", "A->Z", "B->Y", "A->Y"), pairs);

        incoming.add("later", "new", file("later.xlsx"), 2);
        assertNull(mappings.resolve("later").getValue());
        mappings.clear();
        assertEquals(4, incoming.size());
    }

    @Test
    public void returnedMatchesStayUnchangedAfterMoreImportsAndClear() {
        MappingRelations mappings = new MappingRelations();
        mappings.add("A", "X", file("first.xlsx"), 2);
        MappingRelations.Match original = mappings.resolve("A");
        mappings.add("A", "X", file("second.xlsx"), 3);
        MappingRelations.Match extraOrigin = mappings.resolve("A");
        mappings.add("A", "Y", file("third.xlsx"), 4);
        assertTrue(mappings.resolve("A").isAmbiguous());
        mappings.clear();

        assertEquals("X", original.getValue());
        assertEquals(Collections.singletonList("first.xlsx:2"), originLabels(original.getCandidates().get("X")));
        assertEquals(Arrays.asList("first.xlsx:2", "second.xlsx:3"), originLabels(extraOrigin.getCandidates().get("X")));
        assertFalse(original.isAmbiguous());
        assertTrue(mappings.isEmpty());
        assertEquals(0, mappings.size());
        assertTrue(mappings.getSourceFiles().isEmpty());
        assertMiss(mappings.resolve("A"));
        assertMiss(mappings.resolve("X"));
        mappings.forEachPair((source, target) -> fail("清空后仍有关系"));
        mappings.forEachConflict((key, match) -> fail("清空后仍有冲突"));
    }

    @Test
    public void candidateMapsOriginListsAndSourceFilesCannotMutateTheModel() {
        MappingRelations mappings = new MappingRelations();
        mappings.add("A", "X", file("first.xlsx"), 2);
        Map<String, List<MappingRelations.Origin>> candidates = mappings.resolve("A").getCandidates();
        Set<File> sourceFiles = mappings.getSourceFiles();

        assertThrows(UnsupportedOperationException.class, candidates::clear);
        assertThrows(UnsupportedOperationException.class, () -> candidates.get("X").clear());
        assertThrows(UnsupportedOperationException.class, sourceFiles::clear);

        assertEquals("X", mappings.resolve("A").getValue());
        assertEquals(1, mappings.size());
        assertEquals(file("first.xlsx"), candidates.get("X").get(0).getFile());
        assertEquals(2, candidates.get("X").get(0).getRowNumber());
    }

    @Test
    public void normalizesFilePathAliasesWhenDeduplicatingAnOrigin() {
        MappingRelations mappings = new MappingRelations();
        File original = file("first.xlsx");
        File alias = file("sub/../first.xlsx");
        mappings.add("A", "X", original, 2);
        mappings.add("A", "X", alias, 2);

        assertEquals(1, mappings.getSourceFiles().size());
        assertEquals(1, mappings.resolve("A").getCandidates().get("X").size());
        assertEquals(original, mappings.resolve("A").getCandidates().get("X").get(0).getFile());
    }

    @Test
    public void lookupKeepsLiteralWhitespaceAndCaseAndMissesHaveNoCandidate() {
        MappingRelations mappings = new MappingRelations();
        mappings.add(" A ", " X ", file("first.xlsx"), 2);
        mappings.add("A", "Y", file("first.xlsx"), 3);

        assertEquals(" X ", mappings.resolve(" A ").getValue());
        assertEquals("Y", mappings.resolve("A").getValue());
        assertMiss(mappings.resolve("a"));
        assertMiss(mappings.resolve("unknown"));
        assertMiss(mappings.resolve(null));
    }

    @Test
    public void rejectsInvalidOriginRowsWithoutAddingARelation() {
        MappingRelations mappings = new MappingRelations();

        assertThrows(IllegalArgumentException.class, () -> mappings.add("A", "X", file("first.xlsx"), 0));
        assertThrows(IllegalArgumentException.class, () -> mappings.add("A", "X", file("first.xlsx"), -1));

        assertTrue(mappings.isEmpty());
        assertTrue(mappings.getSourceFiles().isEmpty());
        assertMiss(mappings.resolve("A"));
    }

    private File file(String name) {
        return new File(temporary.getRoot(), name);
    }

    private static List<String> originLabels(List<MappingRelations.Origin> origins) {
        List<String> labels = new ArrayList<>();
        for (MappingRelations.Origin origin : origins) {
            labels.add(origin.getFile().getName() + ":" + origin.getRowNumber());
        }
        return labels;
    }

    private static void assertMiss(MappingRelations.Match match) {
        assertFalse(match.isAmbiguous());
        assertFalse(match.isReverse());
        assertNull(match.getValue());
        assertTrue(match.getCandidates().isEmpty());
    }
}
