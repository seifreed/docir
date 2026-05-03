use super::*;
use crate::ole_header::{END_OF_CHAIN, FREE_SECT, SIGNATURE};
use std::collections::HashMap;

fn valid_header_template() -> Vec<u8> {
    let mut data = vec![0u8; 512];
    data[..8].copy_from_slice(&SIGNATURE);
    data[0x1E..0x20].copy_from_slice(&(9u16).to_le_bytes());
    data[0x20..0x22].copy_from_slice(&(6u16).to_le_bytes());
    data[0x2C..0x30].copy_from_slice(&(1u32).to_le_bytes());
    data[0x30..0x34].copy_from_slice(&(0u32).to_le_bytes());
    data[0x38..0x3C].copy_from_slice(&(4096u32).to_le_bytes());
    data[0x3C..0x40].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    data[0x40..0x44].copy_from_slice(&(0u32).to_le_bytes());
    data[0x44..0x48].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    data[0x48..0x4C].copy_from_slice(&(0u32).to_le_bytes());
    for i in 0..109usize {
        let off = 0x4C + i * 4;
        data[off..off + 4].copy_from_slice(&FREE_SECT.to_le_bytes());
    }
    data
}

#[test]
fn ole_signature_detection_and_header_parse() {
    assert!(!is_ole_container(b"not-ole"));
    let header = valid_header_template();
    assert!(is_ole_container(&header));

    let parsed = parse_header(&header).expect("valid header");
    assert_eq!(parsed.sector_size, 512);
    assert_eq!(parsed.mini_sector_size, 64);
    assert_eq!(parsed.num_fat_sectors, 1);
    assert_eq!(parsed.mini_cutoff, 4096);
}

#[test]
fn parse_header_rejects_invalid_signature() {
    let mut bad = vec![0u8; 512];
    bad[..8].copy_from_slice(b"BADHDR!!");
    match parse_header(&bad) {
        Ok(_) => panic!("invalid header should fail"),
        Err(err) => assert!(matches!(err, ParseError::InvalidStructure(_))),
    }
}

#[test]
fn difat_chain_uses_header_entries() {
    let mut header = valid_header_template();
    header[0x4C..0x50].copy_from_slice(&(3u32).to_le_bytes());
    header[0x50..0x54].copy_from_slice(&(7u32).to_le_bytes());

    let parsed = parse_header(&header).expect("header");
    let difat = read_difat_chain(&header, &parsed).expect("difat");
    assert_eq!(difat, vec![3, 7]);
}

#[test]
fn read_sector_and_stream_helpers_handle_bounds_and_chains() {
    let mut data = vec![0u8; 1536];
    data[512..1024].fill(1);
    data[1024..1536].fill(2);

    let s1 = read_sector(&data, 512, 0).expect("sector 0");
    let s2 = read_sector(&data, 512, 1).expect("sector 1");
    assert_eq!(s1[0], 1);
    assert_eq!(s2[0], 2);
    assert!(read_sector(&data, 512, 99).is_err());

    let fat = vec![1, END_OF_CHAIN];
    let stream = read_stream_from_fat(&data, 512, &fat, 0).expect("fat stream");
    assert_eq!(stream.len(), 1024);
    assert_eq!(stream[0], 1);
    assert_eq!(stream[512], 2);

    let mini_stream = b"abcdEFGHijklMNOP".to_vec();
    let mini_fat = vec![1, END_OF_CHAIN];
    let mini = read_stream_from_mini(&mini_stream, 8, &mini_fat, 0, 12).expect("mini stream");
    assert_eq!(mini, b"abcdEFGHijkl".to_vec());
}

#[test]
fn directory_parsing_and_tree_walk_collect_stream_paths() {
    let mut dir = vec![0u8; 128 * 3];

    dir[66] = 5;
    dir[68..72].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[72..76].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[76..80].copy_from_slice(&(1u32).to_le_bytes());
    dir[116..120].copy_from_slice(&(0u32).to_le_bytes());
    dir[120..128].copy_from_slice(&(0u64).to_le_bytes());

    let name1: Vec<u8> = "VBA"
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .chain([0, 0])
        .collect();
    dir[128..128 + name1.len()].copy_from_slice(&name1);
    dir[128 + 64..128 + 66].copy_from_slice(&((name1.len()) as u16).to_le_bytes());
    dir[128 + 66] = 2;
    dir[128 + 68..128 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[128 + 72..128 + 76].copy_from_slice(&(2u32).to_le_bytes());
    dir[128 + 76..128 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[128 + 116..128 + 120].copy_from_slice(&(1u32).to_le_bytes());
    dir[128 + 120..128 + 128].copy_from_slice(&(10u64).to_le_bytes());

    let name2: Vec<u8> = "dir"
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .chain([0, 0])
        .collect();
    dir[256..256 + name2.len()].copy_from_slice(&name2);
    dir[256 + 64..256 + 66].copy_from_slice(&((name2.len()) as u16).to_le_bytes());
    dir[256 + 66] = 2;
    dir[256 + 68..256 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[256 + 72..256 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[256 + 76..256 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[256 + 116..256 + 120].copy_from_slice(&(2u32).to_le_bytes());
    dir[256 + 120..256 + 128].copy_from_slice(&(20u64).to_le_bytes());

    let entries = directory::parse_dir_entries(&dir).expect("dir entries");
    assert_eq!(entries.len(), 3);
    assert_eq!(entries[1].name, "VBA");
    assert_eq!(entries[2].name, "dir");

    let mut out = HashMap::new();
    let mut visited = std::collections::HashSet::new();
    types::walk_siblings(entries[0].child, "", &entries, &mut out, 0, &mut visited);
    assert!(out.contains_key("VBA"));
    assert!(out.contains_key("dir"));

    let metadata = collect_entry_metadata(&entries);
    assert!(metadata.contains_key("Root Entry"));
    assert!(metadata.contains_key("VBA"));
    assert_eq!(metadata["VBA"].entry_type, CfbEntryType::Stream);
}

#[test]
fn collect_directory_slots_marks_normal_orphaned_and_free_entries() {
    let mut dir = vec![0u8; 128 * 4];

    dir[66] = 5;
    dir[68..72].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[72..76].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[76..80].copy_from_slice(&(1u32).to_le_bytes());

    let linked_name: Vec<u8> = "WordDocument"
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .chain([0, 0])
        .collect();
    dir[128..128 + linked_name.len()].copy_from_slice(&linked_name);
    dir[128 + 64..128 + 66].copy_from_slice(&(linked_name.len() as u16).to_le_bytes());
    dir[128 + 66] = 2;
    dir[128 + 68..128 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[128 + 72..128 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[128 + 76..128 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[128 + 116..128 + 120].copy_from_slice(&(1u32).to_le_bytes());
    dir[128 + 120..128 + 128].copy_from_slice(&(3u64).to_le_bytes());

    let orphan_name: Vec<u8> = "Ghost"
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .chain([0, 0])
        .collect();
    dir[256..256 + orphan_name.len()].copy_from_slice(&orphan_name);
    dir[256 + 64..256 + 66].copy_from_slice(&(orphan_name.len() as u16).to_le_bytes());
    dir[256 + 66] = 2;
    dir[256 + 68..256 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[256 + 72..256 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[256 + 76..256 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[256 + 116..256 + 120].copy_from_slice(&(2u32).to_le_bytes());
    dir[256 + 120..256 + 128].copy_from_slice(&(5u64).to_le_bytes());

    dir[384 + 68..384 + 72].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[384 + 72..384 + 76].copy_from_slice(&FREE_SECT.to_le_bytes());
    dir[384 + 76..384 + 80].copy_from_slice(&FREE_SECT.to_le_bytes());

    let entries = directory::parse_dir_entries(&dir).expect("dir entries");
    let slots = collect_directory_slots(&entries);

    assert!(slots
        .iter()
        .any(|slot| slot.path == "WordDocument" && slot.state == CfbDirectoryState::Normal));
    assert!(slots
        .iter()
        .any(|slot| slot.path == "Ghost" && slot.state == CfbDirectoryState::Orphaned));
    assert!(slots
        .iter()
        .any(|slot| slot.entry_index == 3 && slot.state == CfbDirectoryState::Free));
    assert!(slots.iter().any(|slot| slot.path == "WordDocument"
        && slot.name_len_raw as usize == linked_name.len()
        && slot.color_flag_raw == 0
        && slot.left_sibling_raw == FREE_SECT
        && slot.right_sibling_raw == FREE_SECT
        && slot.child_raw == FREE_SECT));
}

#[test]
fn cfb_reader_lists_and_filters_stream_names() {
    let mut streams = HashMap::new();
    streams.insert(
        "VBA/dir".to_string(),
        DirEntry {
            name: "dir".to_string(),
            name_len_raw: 8,
            object_type: 2,
            color_flag: 0,
            left: FREE_SECT,
            right: FREE_SECT,
            child: FREE_SECT,
            start_sector: 0,
            size: 3,
            created_filetime: None,
            modified_filetime: None,
        },
    );
    streams.insert(
        "VBA/Module1".to_string(),
        DirEntry {
            name: "Module1".to_string(),
            name_len_raw: 16,
            object_type: 2,
            color_flag: 0,
            left: FREE_SECT,
            right: FREE_SECT,
            child: FREE_SECT,
            start_sector: 1,
            size: 4,
            created_filetime: None,
            modified_filetime: None,
        },
    );

    let mut entries = HashMap::new();
    entries.insert(
        "Root Entry".to_string(),
        CfbEntryMetadata {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: CfbEntryType::RootStorage,
            object_type_raw: 5,
            size: 0,
            start_sector: END_OF_CHAIN,
            left_sibling: None,
            right_sibling: None,
            child: Some(1),
            created_filetime: None,
            modified_filetime: None,
        },
    );
    entries.insert(
        "VBA/dir".to_string(),
        CfbEntryMetadata {
            entry_index: 1,
            path: "VBA/dir".to_string(),
            entry_type: CfbEntryType::Stream,
            object_type_raw: 2,
            size: 3,
            start_sector: 0,
            left_sibling: None,
            right_sibling: Some(2),
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    );
    entries.insert(
        "VBA/Module1".to_string(),
        CfbEntryMetadata {
            entry_index: 2,
            path: "VBA/Module1".to_string(),
            entry_type: CfbEntryType::Stream,
            object_type_raw: 2,
            size: 4,
            start_sector: 1,
            left_sibling: None,
            right_sibling: None,
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    );

    let cfb = Cfb {
        sector_size: 512,
        mini_sector_size: 64,
        mini_cutoff: 4096,
        num_fat_sectors: 1,
        first_dir_sector: 0,
        first_mini_fat: END_OF_CHAIN,
        num_mini_fat: 0,
        first_difat: END_OF_CHAIN,
        num_difat: 0,
        difat_entry_count: 0,
        fat: vec![END_OF_CHAIN, END_OF_CHAIN],
        mini_fat: Vec::new(),
        root_stream: Vec::new(),
        streams,
        entries,
        directory_slots: Vec::new(),
        data: vec![0u8; 1536],
    };
    let mut reader = CfbReader::new(&cfb);

    assert!(reader.contains("VBA/dir"));
    assert_eq!(reader.file_size("VBA/Module1").expect("size"), 4);
    assert!(reader.list_prefix("VBA/").len() >= 2);
    assert_eq!(reader.list_suffix("dir"), vec!["VBA/dir".to_string()]);
    assert!(reader.read_file("missing").is_err());
}

#[test]
fn cfb_path_normalization_and_read_file_string_error_paths() {
    let mut streams = HashMap::new();
    streams.insert(
        "VBA/Module1".to_string(),
        DirEntry {
            name: "Module1".to_string(),
            name_len_raw: 16,
            object_type: 2,
            color_flag: 0,
            left: FREE_SECT,
            right: FREE_SECT,
            child: FREE_SECT,
            start_sector: 0,
            size: 2,
            created_filetime: None,
            modified_filetime: None,
        },
    );

    let mut entries = HashMap::new();
    entries.insert(
        "Root Entry".to_string(),
        CfbEntryMetadata {
            entry_index: 0,
            path: "Root Entry".to_string(),
            entry_type: CfbEntryType::RootStorage,
            object_type_raw: 5,
            size: 0,
            start_sector: END_OF_CHAIN,
            left_sibling: None,
            right_sibling: None,
            child: Some(1),
            created_filetime: None,
            modified_filetime: None,
        },
    );
    entries.insert(
        "VBA/Module1".to_string(),
        CfbEntryMetadata {
            entry_index: 1,
            path: "VBA/Module1".to_string(),
            entry_type: CfbEntryType::Stream,
            object_type_raw: 2,
            size: 2,
            start_sector: 0,
            left_sibling: None,
            right_sibling: None,
            child: None,
            created_filetime: None,
            modified_filetime: None,
        },
    );

    let mut data = vec![0u8; 1024];
    data[512] = 0xFF;
    data[513] = 0xFE;

    let cfb = Cfb {
        sector_size: 512,
        mini_sector_size: 64,
        mini_cutoff: 4096,
        num_fat_sectors: 1,
        first_dir_sector: 0,
        first_mini_fat: END_OF_CHAIN,
        num_mini_fat: 0,
        first_difat: END_OF_CHAIN,
        num_difat: 0,
        difat_entry_count: 0,
        fat: vec![END_OF_CHAIN],
        mini_fat: Vec::new(),
        root_stream: Vec::new(),
        streams,
        entries,
        directory_slots: Vec::new(),
        data,
    };
    let mut reader = CfbReader::new(&cfb);

    assert!(reader.contains("VBA\\Module1"));
    let err = reader
        .read_file_string("VBA\\Module1")
        .expect_err("invalid UTF-8 should fail");
    assert!(matches!(err, ParseError::Encoding(_)));
}

#[test]
fn read_stream_from_mini_handles_out_of_bounds_and_chain_breaks() {
    let mini_stream = b"abcdEFGH".to_vec();
    let mini_fat = vec![END_OF_CHAIN];
    let out = read_stream_from_mini(&mini_stream, 8, &mini_fat, 10, 4).expect("mini stream");
    assert!(out.is_empty());

    let mini_stream = b"abcdefghijklmnop".to_vec();
    let out = read_stream_from_mini(&mini_stream, 8, &mini_fat, 0, 12).expect("mini stream");
    assert_eq!(out, b"abcdefgh".to_vec());
}

#[test]
fn parse_dir_entries_reads_optional_filetime_metadata() {
    let mut dir = vec![0u8; 128 * 2];
    dir[66] = 5;
    dir[76..80].copy_from_slice(&(1u32).to_le_bytes());

    let name1: Vec<u8> = "Storage"
        .encode_utf16()
        .flat_map(|u| u.to_le_bytes())
        .chain([0, 0])
        .collect();
    dir[128..128 + name1.len()].copy_from_slice(&name1);
    dir[128 + 64..128 + 66].copy_from_slice(&((name1.len()) as u16).to_le_bytes());
    dir[128 + 66] = 1;
    dir[128 + 100..128 + 108].copy_from_slice(&(123456u64).to_le_bytes());
    dir[128 + 108..128 + 116].copy_from_slice(&(654321u64).to_le_bytes());

    let entries = directory::parse_dir_entries(&dir).expect("dir entries");
    assert_eq!(entries[1].created_filetime, Some(123456));
    assert_eq!(entries[1].modified_filetime, Some(654321));
}
