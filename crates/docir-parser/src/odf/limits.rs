use crate::error::ParseError;
use crate::parser::ParserConfig;
use std::cell::Cell as StdCell;
use std::sync::atomic::{AtomicU64, Ordering};

pub(super) trait OdfLimitCounter {
    fn fast_mode(&self) -> bool;
    fn sample_rows(&self) -> u32;
    fn sample_cols(&self) -> u32;
    fn bump_cells(&self, add: u64) -> Result<(), ParseError>;
    fn bump_rows(&self, add: u64) -> Result<(), ParseError>;
    fn bump_paragraphs(&self, add: u64) -> Result<(), ParseError>;
}

pub(super) struct OdfLimits {
    fast_mode: bool,
    sample_rows: u32,
    sample_cols: u32,
    max_cells: Option<u64>,
    max_rows: Option<u64>,
    max_paragraphs: Option<u64>,
    cells: StdCell<u64>,
    rows: StdCell<u64>,
    paragraphs: StdCell<u64>,
}

impl OdfLimits {
    pub(super) fn new(config: &ParserConfig, fast_mode: bool) -> Self {
        Self {
            fast_mode,
            sample_rows: config.odf.fast_sample_rows,
            sample_cols: config.odf.fast_sample_cols,
            max_cells: config.odf.max_cells,
            max_rows: config.odf.max_rows,
            max_paragraphs: config.odf.max_paragraphs,
            cells: StdCell::new(0),
            rows: StdCell::new(0),
            paragraphs: StdCell::new(0),
        }
    }

    fn bump(
        counter: &StdCell<u64>,
        add: u64,
        limit: Option<u64>,
        label: &str,
    ) -> Result<(), ParseError> {
        let next = counter.get().saturating_add(add);
        counter.set(next);
        if let Some(max) = limit {
            if next > max {
                return Err(ParseError::ResourceLimit(format!(
                    "ODF max {} exceeded: {} (max: {})",
                    label, next, max
                )));
            }
        }
        Ok(())
    }
}

impl OdfLimitCounter for OdfLimits {
    fn fast_mode(&self) -> bool {
        self.fast_mode
    }

    fn sample_rows(&self) -> u32 {
        self.sample_rows
    }

    fn sample_cols(&self) -> u32 {
        self.sample_cols
    }

    fn bump_cells(&self, add: u64) -> Result<(), ParseError> {
        Self::bump(&self.cells, add, self.max_cells, "cells")
    }

    fn bump_rows(&self, add: u64) -> Result<(), ParseError> {
        Self::bump(&self.rows, add, self.max_rows, "rows")
    }

    fn bump_paragraphs(&self, add: u64) -> Result<(), ParseError> {
        Self::bump(&self.paragraphs, add, self.max_paragraphs, "paragraphs")
    }
}

pub(super) struct OdfAtomicLimits {
    fast_mode: bool,
    sample_rows: u32,
    sample_cols: u32,
    max_cells: Option<u64>,
    max_rows: Option<u64>,
    max_paragraphs: Option<u64>,
    cells: AtomicU64,
    rows: AtomicU64,
    paragraphs: AtomicU64,
}

impl OdfAtomicLimits {
    pub(super) fn new(config: &ParserConfig, fast_mode: bool) -> Self {
        Self {
            fast_mode,
            sample_rows: config.odf.fast_sample_rows,
            sample_cols: config.odf.fast_sample_cols,
            max_cells: config.odf.max_cells,
            max_rows: config.odf.max_rows,
            max_paragraphs: config.odf.max_paragraphs,
            cells: AtomicU64::new(0),
            rows: AtomicU64::new(0),
            paragraphs: AtomicU64::new(0),
        }
    }

    fn bump(
        counter: &AtomicU64,
        add: u64,
        limit: Option<u64>,
        label: &str,
    ) -> Result<(), ParseError> {
        let next = counter
            .fetch_add(add, Ordering::Relaxed)
            .saturating_add(add);
        if let Some(max) = limit {
            if next > max {
                return Err(ParseError::ResourceLimit(format!(
                    "ODF max {} exceeded: {} (max: {})",
                    label, next, max
                )));
            }
        }
        Ok(())
    }
}

impl OdfLimitCounter for OdfAtomicLimits {
    fn fast_mode(&self) -> bool {
        self.fast_mode
    }

    fn sample_rows(&self) -> u32 {
        self.sample_rows
    }

    fn sample_cols(&self) -> u32 {
        self.sample_cols
    }

    fn bump_cells(&self, add: u64) -> Result<(), ParseError> {
        Self::bump(&self.cells, add, self.max_cells, "cells")
    }

    fn bump_rows(&self, add: u64) -> Result<(), ParseError> {
        Self::bump(&self.rows, add, self.max_rows, "rows")
    }

    fn bump_paragraphs(&self, add: u64) -> Result<(), ParseError> {
        Self::bump(&self.paragraphs, add, self.max_paragraphs, "paragraphs")
    }
}
