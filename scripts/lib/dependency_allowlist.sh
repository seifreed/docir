#!/usr/bin/env bash
set -euo pipefail

is_allowed_dependency() {
  local crate_name="$1"
  local dep="$2"

  case "$crate_name" in
    docir-core)
      case "$dep" in
        serde|thiserror)
          return 0
          ;;
      esac
      ;;
    docir-parser)
      case "$dep" in
        docir-core|docir-security|zip|quick-xml|encoding_rs|flate2|calamine|sha2|sha1|pbkdf2|base64|aes|cbc|log|serde|thiserror)
          return 0
          ;;
      esac
      ;;
    docir-app)
      case "$dep" in
        docir-core|docir-parser|docir-security|docir-serialization|docir-diff|docir-rules|serde|thiserror)
          return 0
          ;;
      esac
      ;;
    docir-security)
      case "$dep" in
        docir-core|sha2|log|thiserror)
          return 0
          ;;
      esac
      ;;
    docir-serialization)
      case "$dep" in
        docir-core|serde|serde_json|thiserror)
          return 0
          ;;
      esac
      ;;
    docir-rules)
      case "$dep" in
        docir-core|serde)
          return 0
          ;;
      esac
      ;;
    docir-diff)
      case "$dep" in
        docir-core|serde|serde_json|sha2|thiserror)
          return 0
          ;;
      esac
      ;;
    docir-cli)
      case "$dep" in
        docir-core|docir-app|clap|anyhow|env_logger|log|serde_json|serde)
          return 0
          ;;
      esac
      ;;
    docir-python)
      case "$dep" in
        docir-core|docir-app|pyo3|anyhow|serde|serde_json)
          return 0
          ;;
      esac
      ;;
    *)
      return 0
      ;;
  esac

  return 1
}
