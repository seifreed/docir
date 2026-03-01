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

#[cfg(test)]
mod tests {
    use super::*;
    use docir_core::ir::IRNode;

    #[test]
    fn finalize_object_uses_name_fallback_and_sets_size() {
        let mut store = IrStore::new();
        let object = ObjectContext {
            class_name: Some("Word.Document.12".to_string()),
            object_name: None,
            data_hex_len: 12,
            ..ObjectContext::default()
        };

        let id = finalize_object(object, &mut store).expect("object id");
        let Some(IRNode::OleObject(ole)) = store.get(id) else {
            panic!("expected OLE object");
        };

        assert_eq!(ole.name.as_deref(), Some("Word.Document.12"));
        assert_eq!(ole.prog_id.as_deref(), Some("Word.Document.12"));
        assert_eq!(ole.size_bytes, 6);
    }

    #[test]
    fn finalize_picture_sets_content_type_for_media_variants() {
        let mut store = IrStore::new();

        let audio = ObjectContext {
            data_hex_len: 20,
            media_type: Some(MediaType::Audio),
            ..ObjectContext::default()
        };
        let audio_id = finalize_picture(audio, &mut store).expect("audio id");
        let Some(IRNode::MediaAsset(asset)) = store.get(audio_id) else {
            panic!("expected media asset");
        };
        assert_eq!(asset.media_type, MediaType::Audio);
        assert_eq!(asset.content_type.as_deref(), Some("audio/rtf"));
        assert_eq!(asset.size_bytes, 10);
        assert_eq!(
            asset.span.as_ref().map(|s| s.file_path.as_str()),
            Some("rtf")
        );

        let other = ObjectContext {
            data_hex_len: 2,
            media_type: Some(MediaType::Other),
            ..ObjectContext::default()
        };
        let other_id = finalize_picture(other, &mut store).expect("other id");
        let Some(IRNode::MediaAsset(asset)) = store.get(other_id) else {
            panic!("expected media asset");
        };
        assert_eq!(asset.media_type, MediaType::Other);
        assert_eq!(asset.content_type, None);
    }

    #[test]
    fn finalize_picture_defaults_to_image_when_media_type_missing() {
        let mut store = IrStore::new();
        let object = ObjectContext {
            data_hex_len: 8,
            ..ObjectContext::default()
        };
        let id = finalize_picture(object, &mut store).expect("media id");
        let Some(IRNode::MediaAsset(asset)) = store.get(id) else {
            panic!("expected media asset");
        };
        assert_eq!(asset.media_type, MediaType::Image);
        assert_eq!(asset.content_type, None);
        assert_eq!(asset.size_bytes, 4);
    }
}
