use serde::Serialize;

/// Generic bucket-count pair for categorical frequency distributions.
///
/// Replaces the former per-category `DirectoryXxxCount` and sector count
/// structs, all of which carried the same `{ label, count }` shape under
/// different names.
#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
pub struct BucketCount {
    pub bucket: String,
    pub count: usize,
}

impl BucketCount {
    pub fn new(bucket: impl Into<String>, count: usize) -> Self {
        Self {
            bucket: bucket.into(),
            count,
        }
    }
}
