use crate::io_support::read_bounded_file;
use crate::{AppResult, ParserConfig};
use docir_parser::ole::{Cfb, CfbEntryType};
use serde::Serialize;
use std::path::Path;

/// Dedicated timestamp listing for CFB storages and streams.
#[derive(Debug, Clone, Serialize)]
pub struct TimeListing {
    pub container: String,
    pub entry_count: usize,
    pub entries: Vec<TimeEntry>,
}

/// One timestamp-bearing CFB directory entry.
#[derive(Debug, Clone, Serialize)]
pub struct TimeEntry {
    pub path: String,
    pub entry_type: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created_filetime: Option<u64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified_filetime: Option<u64>,
}

/// Lists FILETIMEs from a legacy CFB/OLE container on disk.
pub fn list_times_path<P: AsRef<Path>>(path: P, config: &ParserConfig) -> AppResult<TimeListing> {
    let bytes = read_bounded_file(path, config.max_input_size)?;
    list_times_bytes(&bytes)
}

/// Lists FILETIMEs from raw CFB bytes.
pub fn list_times_bytes(data: &[u8]) -> AppResult<TimeListing> {
    let cfb = Cfb::parse(data.to_vec())?;
    let mut entries: Vec<TimeEntry> = cfb
        .list_entries()
        .into_iter()
        .map(|entry| TimeEntry {
            path: entry.path,
            entry_type: match entry.entry_type {
                CfbEntryType::RootStorage => "root-storage".to_string(),
                CfbEntryType::Storage => "storage".to_string(),
                CfbEntryType::Stream => "stream".to_string(),
            },
            created_filetime: entry.created_filetime,
            modified_filetime: entry.modified_filetime,
        })
        .collect();
    entries.sort_by(|left, right| left.path.cmp(&right.path));
    Ok(TimeListing {
        container: "cfb-ole".to_string(),
        entry_count: entries.len(),
        entries,
    })
}

#[cfg(test)]
mod tests {
    use super::list_times_bytes;
    use crate::test_support::build_test_cfb_with_times;

    #[test]
    fn list_times_reads_cfb_entry_timestamps() {
        let created = 132_537_600_000_000_000u64;
        let modified = 132_537_600_123_456_789u64;
        let listing = list_times_bytes(&build_test_cfb_with_times(
            &[("WordDocument", b"doc"), ("VBA/PROJECT", b"meta")],
            &[
                ("WordDocument", created, modified),
                ("VBA", created + 1, modified + 1),
            ],
        ))
        .expect("listing");

        let word = listing
            .entries
            .iter()
            .find(|entry| entry.path == "WordDocument")
            .expect("word");
        assert_eq!(word.entry_type, "stream");
        assert_eq!(word.created_filetime, Some(created));
        assert_eq!(word.modified_filetime, Some(modified));

        let vba = listing
            .entries
            .iter()
            .find(|entry| entry.path == "VBA")
            .expect("vba storage");
        assert_eq!(vba.entry_type, "storage");
        assert_eq!(vba.created_filetime, Some(created + 1));
        assert_eq!(vba.modified_filetime, Some(modified + 1));
    }
}
