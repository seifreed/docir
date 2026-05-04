use crate::error::ParseError;
use crate::xml_utils::lossy_attr_value;
use crate::xml_utils::{read_event, reader_from_str};
use docir_core::ir::{PeoplePart, PersonEntry};
use docir_core::types::SourceSpan;
use quick_xml::events::Event;

/// Public API entrypoint: parse_people_part.
pub fn parse_people_part(xml: &str, path: &str) -> Result<PeoplePart, ParseError> {
    parse_people_part_impl(xml, path)
}

fn parse_people_part_impl(xml: &str, path: &str) -> Result<PeoplePart, ParseError> {
    let mut reader = reader_from_str(xml);

    let mut people = PeoplePart::new();
    people.span = Some(SourceSpan::new(path));

    let mut buf = Vec::new();
    loop {
        match read_event(&mut reader, &mut buf, path)? {
            Event::Start(e) | Event::Empty(e) => {
                if e.name().as_ref().ends_with(b"person") {
                    let mut entry = PersonEntry {
                        person_id: None,
                        user_id: None,
                        display_name: None,
                        initials: None,
                    };
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let key = match key.iter().rposition(|b| *b == b':') {
                            Some(pos) => &key[pos + 1..],
                            None => key,
                        };
                        let val = lossy_attr_value(&attr).to_string();
                        match key {
                            b"id" => entry.person_id = Some(val),
                            b"userId" | b"userID" => entry.user_id = Some(val),
                            b"displayName" | b"displayname" => entry.display_name = Some(val),
                            b"initials" => entry.initials = Some(val),
                            _ => {}
                        }
                    }
                    if entry.person_id.is_some()
                        || entry.user_id.is_some()
                        || entry.display_name.is_some()
                        || entry.initials.is_some()
                    {
                        people.people.push(entry);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(people)
}
