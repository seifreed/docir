use super::inspect_sectors_bytes;
use crate::inspect_directory_bytes;
use crate::test_support::{
    build_test_cfb, patch_test_cfb_directory_entry, patch_test_cfb_fat_entry,
    patch_test_cfb_header_u32, TestCfbDirectoryPatch,
};

#[test]
fn inspect_sectors_reads_fat_and_stream_chains() {
    let inspection = inspect_sectors_bytes(&build_test_cfb(&[
        ("WordDocument", b"doc"),
        ("VBA/PROJECT", b"meta"),
    ]))
    .expect("inspection");

    assert_eq!(inspection.container, "cfb-ole");
    assert_eq!(inspection.sector_size, 512);
    assert_eq!(inspection.sector_count, 4);
    assert_eq!(inspection.num_fat_sectors, 1);
    assert_eq!(inspection.first_dir_sector, 2);
    assert_eq!(inspection.num_difat, 0);
    assert_eq!(inspection.difat_entry_count, 109);
    assert_eq!(inspection.fat_reserved_count, 1);
    assert_eq!(inspection.occupied_fat_entries, 128);
    assert!(inspection.anomalies.is_empty());
    assert!(inspection
        .role_counts
        .iter()
        .any(|entry| entry.role == "stream-data" && entry.count >= 1));
    assert!(inspection
        .role_counts
        .iter()
        .any(|entry| entry.role == "special:directory" && entry.count >= 1));
    assert!(inspection
        .sector_overview
        .iter()
        .any(|entry| entry.sector == 0
            && entry.role == "stream-data"
            && entry.fat_value == "end-of-chain"
            && entry.special_roles.is_empty()
            && entry.owners.len() == 1
            && entry.owners[0].path == "WordDocument"
            && entry.owners[0].index_in_chain == 0
            && entry.owners[0].is_terminal));
    assert!(inspection
        .sector_overview
        .iter()
        .any(|entry| entry.sector == 2 && entry.special_roles.contains(&"directory".to_string())));
    assert!(inspection
        .streams
        .iter()
        .any(|stream| stream.path == "WordDocument"
            && stream.sector_chain == vec![0]
            && stream.chain_state == "complete"
            && stream.chain_terminal_raw == u32::MAX - 1
            && stream.chain_terminal == "end-of-chain"
            && stream.start_sector == 0
            && stream.logical_root == "WordDocument"
            && stream.size_bytes == 3
            && stream.expected_chain_len == 1
            && stream.chain_steps.len() == 1
            && stream.chain_steps[0].sector == 0
            && stream.chain_steps[0].next_raw == u32::MAX - 1));
    assert!(inspection
        .streams
        .iter()
        .any(|stream| stream.path == "VBA/PROJECT" && stream.sector_chain == vec![1]));
}

#[test]
fn inspect_sectors_detects_shared_sector_claims() {
    let base = build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]);
    let inspection = inspect_directory_bytes(&base).expect("directory");
    let vba_entry = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "VBA/PROJECT")
        .expect("vba entry");
    let patched = patch_test_cfb_directory_entry(
        &base,
        vba_entry.entry_index,
        TestCfbDirectoryPatch {
            start_sector: Some(0),
            ..Default::default()
        },
    );

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    assert!(inspection
        .shared_sector_claims
        .iter()
        .any(|claim| claim.sector == 0 && claim.owners.len() >= 2));
    assert!(inspection
        .anomalies
        .iter()
        .any(|value| value.kind == "shared-sector"
            && value.message.contains("claimed by multiple streams")
            && value.severity == "high"));
}

#[test]
fn inspect_sectors_detects_shared_chain_overlaps() {
    let base = build_test_cfb(&[
        ("WordDocument", &[0u8; 1025]),
        ("VBA/PROJECT", &[1u8; 1025]),
    ]);
    let inspection = inspect_directory_bytes(&base).expect("directory");
    let vba_entry = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "VBA/PROJECT")
        .expect("vba entry");
    let patched = patch_test_cfb_directory_entry(
        &base,
        vba_entry.entry_index,
        TestCfbDirectoryPatch {
            start_sector: Some(0),
            ..Default::default()
        },
    );

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    assert!(inspection
        .shared_chain_overlaps
        .iter()
        .any(|overlap| overlap.sectors.len() >= 2));
}

#[test]
fn inspect_sectors_summarizes_truncated_chains_by_logical_root() {
    let base = build_test_cfb(&[("WordDocument", &[0u8; 600]), ("VBA/PROJECT", b"meta")]);
    let patched = patch_test_cfb_fat_entry(&base, 0, u32::MAX - 1);

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    assert!(inspection
        .truncated_chain_counts
        .iter()
        .any(|entry| entry.bucket == "fat:WordDocument" && entry.count == 1));
    assert!(inspection
        .streams
        .iter()
        .any(|stream| stream.path == "WordDocument" && stream.chain_state == "truncated"));
}

#[test]
fn inspect_sectors_summarizes_structural_incoherence() {
    let base = build_test_cfb(&[("WordDocument", b"doc")]);
    let patched = patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&base, 0x3C, 0), 0x40, 1);

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    assert!(inspection
        .structural_incoherence_counts
        .iter()
        .any(|entry| entry.bucket == "mini-fat-without-consumers"
            && entry.severity == "medium"
            && entry.count == 1));
}

#[test]
fn inspect_sectors_detects_start_sector_mismatch() {
    let base = build_test_cfb(&[("WordDocument", &[0u8; 600])]);
    let inspection = inspect_directory_bytes(&base).expect("directory");
    let word_entry = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "WordDocument")
        .expect("word entry");
    let patched = patch_test_cfb_fat_entry(&base, word_entry.start_sector, 0xFFFF_FFFF);

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    assert!(inspection
        .structural_incoherence_counts
        .iter()
        .any(
            |entry| entry.bucket.contains("start-sector-mismatch:WordDocument")
                && entry.severity == "high"
        ));
}

#[test]
fn inspect_sectors_detects_start_sector_reuse_and_chain_health() {
    let base = build_test_cfb(&[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")]);
    let inspection = inspect_directory_bytes(&base).expect("directory");
    let vba_entry = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "VBA/PROJECT")
        .expect("vba entry");
    let patched = patch_test_cfb_directory_entry(
        &base,
        vba_entry.entry_index,
        TestCfbDirectoryPatch {
            start_sector: Some(0),
            ..Default::default()
        },
    );

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    assert!(inspection
        .start_sector_reuse
        .iter()
        .any(|entry| entry.sector == 0 && entry.owners.len() >= 2));
    assert!(inspection
        .chain_health_by_root
        .iter()
        .any(|entry| entry.bucket.contains("health:start-reused:root:") && entry.count >= 1));
}

#[test]
fn inspect_sectors_assigns_high_risk_to_shared_and_invalid_start_streams() {
    let base = build_test_cfb(&[
        ("WordDocument", &[0u8; 1025]),
        ("VBA/PROJECT", &[1u8; 1025]),
    ]);
    let inspection = inspect_directory_bytes(&base).expect("directory");
    let word_entry = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "WordDocument")
        .expect("word entry");
    let vba_entry = inspection
        .entries
        .iter()
        .find(|entry| entry.path == "VBA/PROJECT")
        .expect("vba entry");
    let patched = patch_test_cfb_directory_entry(
        &patch_test_cfb_directory_entry(
            &base,
            vba_entry.entry_index,
            TestCfbDirectoryPatch {
                start_sector: Some(word_entry.start_sector),
                ..Default::default()
            },
        ),
        word_entry.entry_index,
        TestCfbDirectoryPatch {
            start_sector: Some(99),
            ..Default::default()
        },
    );

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    let word = inspection
        .streams
        .iter()
        .find(|stream| stream.path == "WordDocument")
        .expect("word stream");
    let vba = inspection
        .streams
        .iter()
        .find(|stream| stream.path == "VBA/PROJECT")
        .expect("vba stream");
    assert_eq!(word.stream_health, "invalid-start");
    assert_eq!(word.stream_risk, "high");
    assert_eq!(vba.stream_health, "shared");
    assert_eq!(vba.stream_risk, "high");
    assert!(inspection.shared_chain_overlaps.iter().any(|overlap| {
        overlap.owners.iter().any(|owner| owner == "WordDocument") && overlap.severity == "high"
    }));
}

#[test]
fn inspect_sectors_reports_mini_fat_incoherence_with_severity() {
    let base = build_test_cfb(&[("Tiny", &[0u8; 32])]);
    let patched = patch_test_cfb_header_u32(&patch_test_cfb_header_u32(&base, 0x3C, 0), 0x40, 1);

    let inspection = inspect_sectors_bytes(&patched).expect("inspection");
    assert!(inspection
        .anomalies
        .iter()
        .any(|entry| entry.kind == "mini-fat-without-consumers" && entry.severity == "medium"));
}
