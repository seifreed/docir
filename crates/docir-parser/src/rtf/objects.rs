use docir_core::ir::{MediaAsset, MediaType};
use docir_core::security::OleObject;
use docir_core::types::{NodeId, SourceSpan};
use docir_core::visitor::IrStore;

#[derive(Debug, Default)]
pub(crate) struct ObjectContext {
    pub(crate) class_name: Option<String>,
    pub(crate) object_name: Option<String>,
    pub(crate) data_hex_len: usize,
    pub(crate) media_type: Option<MediaType>,
    pub(crate) pic_width: Option<u32>,
    pub(crate) pic_height: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ObjectTextTarget {
    Class,
    Name,
}

pub(crate) fn finalize_object(object: ObjectContext, store: &mut IrStore) -> Option<NodeId> {
    let mut ole = OleObject::new();
    ole.name = object
        .object_name
        .clone()
        .or_else(|| object.class_name.clone());
    ole.prog_id = object.class_name.clone();
    ole.size_bytes = (object.data_hex_len / 2) as u64;
    let ole_id = ole.id;
    store.insert(docir_core::ir::IRNode::OleObject(ole));
    Some(ole_id)
}

pub(crate) fn finalize_picture(object: ObjectContext, store: &mut IrStore) -> Option<NodeId> {
    let size = (object.data_hex_len / 2) as u64;
    let media_type = object.media_type.unwrap_or(MediaType::Image);
    let mut asset = MediaAsset::new("rtf/pict", media_type, size);
    if let Some(mt) = object.media_type {
        asset.content_type = match mt {
            MediaType::Image => Some("image/rtf".to_string()),
            MediaType::Audio => Some("audio/rtf".to_string()),
            MediaType::Video => Some("video/rtf".to_string()),
            MediaType::Other => None,
        };
    }
    asset.span = Some(SourceSpan::new("rtf"));
    let id = asset.id;
    store.insert(docir_core::ir::IRNode::MediaAsset(asset));
    Some(id)
}
