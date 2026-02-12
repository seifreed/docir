//! Visitor pattern for IR traversal.
//!
//! This module provides traits and utilities for traversing the IR tree
//! in various orders (pre-order, post-order) and performing operations
//! on nodes.

use crate::error::CoreError;
use crate::ir::node_list::for_each_ir_node;
use crate::ir::DigitalSignature as IrDigitalSignature;
use crate::ir::*;
use crate::security::*;

mod store;
mod visitors;
mod walk;

type DigitalSignature = IrDigitalSignature;

/// Result type for visitor operations.
pub type VisitorResult<T> = Result<T, CoreError>;

/// Control flow for visitor traversal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VisitControl {
    /// Continue traversal normally.
    Continue,
    /// Skip children of current node.
    SkipChildren,
    /// Stop traversal entirely.
    Stop,
}

macro_rules! define_visit_defaults {
    ($variant:ident, $ty:ident, $method:ident) => {
        fn $method(&mut self, _node: &$ty) -> VisitorResult<VisitControl> {
            Ok(VisitControl::Continue)
        }
    };
}

/// Trait for immutable IR traversal.
///
/// Implement this trait to perform read-only operations on the IR tree.
/// Default implementations return `Continue` for all node types.
pub trait IrVisitor {
    for_each_ir_node!(define_visit_defaults, ;);
}

pub use store::IrStore;
pub use visitors::{NodeCounter, TextCollector};
pub use walk::PreOrderWalker;
