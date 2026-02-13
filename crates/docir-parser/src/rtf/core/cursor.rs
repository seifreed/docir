pub(crate) struct RtfCursor<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> RtfCursor<'a> {
    pub(crate) fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }

    pub(super) fn peek(&self) -> Option<u8> {
        self.data.get(self.pos).copied()
    }

    pub(super) fn next(&mut self) -> Option<u8> {
        let b = self.data.get(self.pos).copied();
        if b.is_some() {
            self.pos += 1;
        }
        b
    }

    pub(super) fn is_eof(&self) -> bool {
        self.pos >= self.data.len()
    }
}

pub(crate) fn is_rtf_bytes(data: &[u8]) -> bool {
    data.starts_with(b"{\\rtf") || data.starts_with(b"{\rtf")
}
