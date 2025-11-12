use anyhow::{Context, Result};
use crossterm::event::{KeyCode, KeyModifiers};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

/// Application configuration
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct Config {
    pub theme: ThemeConfig,
    pub ui: UiConfig,
    pub keybindings: KeybindingsConfig,
}

/// Theme configuration
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct ThemeConfig {
    /// Default theme to use on startup
    pub default: String,
}

/// UI configuration
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct UiConfig {
    /// Default maximum rows to display in non-interactive mode
    pub max_rows: usize,
    /// Default maximum column width
    pub column_width: usize,
}

/// Keybindings configuration
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(default)]
pub struct KeybindingsConfig {
    /// Keybinding profile: "default", "vim", or "custom"
    pub profile: String,
    /// Custom keybindings (overrides profile)
    #[serde(skip_serializing_if = "HashMap::is_empty")]
    pub custom: HashMap<String, String>,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            theme: ThemeConfig::default(),
            ui: UiConfig::default(),
            keybindings: KeybindingsConfig::default(),
        }
    }
}

impl Default for ThemeConfig {
    fn default() -> Self {
        Self {
            default: "Default".to_string(),
        }
    }
}

impl Default for UiConfig {
    fn default() -> Self {
        Self {
            max_rows: 50,
            column_width: 30,
        }
    }
}

impl Default for KeybindingsConfig {
    fn default() -> Self {
        Self {
            profile: "default".to_string(),
            custom: HashMap::new(),
        }
    }
}

impl Config {
    /// Load configuration from XDG config directory or custom path
    pub fn load(custom_path: Option<PathBuf>) -> Result<Self> {
        let config_path = if let Some(path) = custom_path {
            path
        } else {
            Self::default_config_path()?
        };

        if !config_path.exists() {
            // No config file, return defaults
            return Ok(Self::default());
        }

        let config_str = fs::read_to_string(&config_path)
            .with_context(|| format!("Failed to read config file: {}", config_path.display()))?;

        let config: Config = toml::from_str(&config_str)
            .with_context(|| format!("Failed to parse config file: {}", config_path.display()))?;

        Ok(config)
    }

    /// Get the default config file path ($XDG_CONFIG_HOME/xleak/config.toml)
    pub fn default_config_path() -> Result<PathBuf> {
        let config_dir = dirs::config_dir()
            .context("Failed to determine config directory")?
            .join("xleak");

        Ok(config_dir.join("config.toml"))
    }

    /// Create an example config file at the default location
    pub fn create_example() -> Result<PathBuf> {
        let config_path = Self::default_config_path()?;
        let config_dir = config_path
            .parent()
            .context("Failed to get config directory")?;

        // Create config directory if it doesn't exist
        fs::create_dir_all(config_dir)
            .with_context(|| format!("Failed to create config directory: {}", config_dir.display()))?;

        // Generate example config
        let example = Self::example_toml();
        fs::write(&config_path, example)
            .with_context(|| format!("Failed to write example config: {}", config_path.display()))?;

        Ok(config_path)
    }

    /// Generate example TOML config
    fn example_toml() -> String {
        r#"# xleak configuration file
# Location: $XDG_CONFIG_HOME/xleak/config.toml (usually ~/.config/xleak/config.toml)

[theme]
# Default theme to use on startup
# Options: "Default", "Dracula", "Solarized Dark", "Solarized Light", "GitHub Dark", "Nord"
default = "Default"

[ui]
# Default maximum rows to display in non-interactive mode (0 = all)
max_rows = 50
# Default maximum column width in characters
column_width = 30

[keybindings]
# Keybinding profile: "default" or "vim"
profile = "default"

# Custom keybindings (optional - overrides profile)
# Uncomment and modify to customize individual keys
# [keybindings.custom]
# quit = "q"
# help = "?"
# theme_toggle = "t"
# search = "/"
# next_match = "n"
# prev_match = "N"
# copy_cell = "c"
# copy_row = "C"
# jump = "Ctrl+g"
# show_cell_detail = "Enter"

# VIM-style navigation (when profile = "vim")
# up = "k"
# down = "j"
# left = "h"
# right = "l"
# page_up = "Ctrl+u"
# page_down = "Ctrl+d"
# jump_to_top = "g"
# jump_to_bottom = "G"
# jump_to_row_start = "0"
# jump_to_row_end = "$"
"#.to_string()
    }

    /// Get keybinding for an action based on profile and custom overrides
    pub fn get_keybinding(&self, action: &str) -> Option<(KeyCode, KeyModifiers)> {
        // Check custom bindings first
        if let Some(key_str) = self.keybindings.custom.get(action) {
            return parse_key_string(key_str);
        }

        // Fall back to profile defaults
        match self.keybindings.profile.as_str() {
            "vim" => get_vim_keybinding(action),
            _ => get_default_keybinding(action),
        }
    }
}

/// Parse a key string like "q", "Ctrl+g", "Enter" into KeyCode and KeyModifiers
fn parse_key_string(s: &str) -> Option<(KeyCode, KeyModifiers)> {
    let parts: Vec<&str> = s.split('+').collect();
    let mut modifiers = KeyModifiers::empty();
    let key_part = parts.last()?;

    // Parse modifiers
    for part in &parts[..parts.len() - 1] {
        match part.to_lowercase().as_str() {
            "ctrl" | "control" => modifiers |= KeyModifiers::CONTROL,
            "alt" => modifiers |= KeyModifiers::ALT,
            "shift" => modifiers |= KeyModifiers::SHIFT,
            _ => return None,
        }
    }

    // Parse key code
    let code = match key_part.to_lowercase().as_str() {
        "enter" => KeyCode::Enter,
        "esc" | "escape" => KeyCode::Esc,
        "tab" => KeyCode::Tab,
        "backtab" => KeyCode::BackTab,
        "backspace" => KeyCode::Backspace,
        "delete" | "del" => KeyCode::Delete,
        "insert" | "ins" => KeyCode::Insert,
        "home" => KeyCode::Home,
        "end" => KeyCode::End,
        "pageup" | "pgup" => KeyCode::PageUp,
        "pagedown" | "pgdn" => KeyCode::PageDown,
        "up" => KeyCode::Up,
        "down" => KeyCode::Down,
        "left" => KeyCode::Left,
        "right" => KeyCode::Right,
        s if s.len() == 1 => KeyCode::Char(s.chars().next()?),
        _ => return None,
    };

    Some((code, modifiers))
}

/// Get default keybinding for an action
fn get_default_keybinding(action: &str) -> Option<(KeyCode, KeyModifiers)> {
    let binding = match action {
        "quit" => ("q", KeyModifiers::empty()),
        "help" => ("?", KeyModifiers::empty()),
        "theme_toggle" => ("t", KeyModifiers::empty()),
        "search" => ("/", KeyModifiers::empty()),
        "next_match" => ("n", KeyModifiers::empty()),
        "prev_match" => ("N", KeyModifiers::SHIFT),
        "copy_cell" => ("c", KeyModifiers::empty()),
        "copy_row" => ("C", KeyModifiers::SHIFT),
        "jump" => ("g", KeyModifiers::CONTROL),
        "show_cell_detail" => ("Enter", KeyModifiers::empty()),
        "next_sheet" => ("Tab", KeyModifiers::empty()),
        "prev_sheet" => ("Tab", KeyModifiers::SHIFT),
        "up" => ("Up", KeyModifiers::empty()),
        "down" => ("Down", KeyModifiers::empty()),
        "left" => ("Left", KeyModifiers::empty()),
        "right" => ("Right", KeyModifiers::empty()),
        "page_up" => ("PageUp", KeyModifiers::empty()),
        "page_down" => ("PageDown", KeyModifiers::empty()),
        "jump_to_top" => ("Home", KeyModifiers::CONTROL),
        "jump_to_bottom" => ("End", KeyModifiers::CONTROL),
        "jump_to_row_start" => ("Home", KeyModifiers::empty()),
        "jump_to_row_end" => ("End", KeyModifiers::empty()),
        _ => return None,
    };

    parse_key_string(binding.0).map(|(code, _)| (code, binding.1))
}

/// Get VIM-style keybinding for an action
fn get_vim_keybinding(action: &str) -> Option<(KeyCode, KeyModifiers)> {
    let binding = match action {
        // VIM navigation
        "up" => ("k", KeyModifiers::empty()),
        "down" => ("j", KeyModifiers::empty()),
        "left" => ("h", KeyModifiers::empty()),
        "right" => ("l", KeyModifiers::empty()),
        "page_up" => ("u", KeyModifiers::CONTROL),
        "page_down" => ("d", KeyModifiers::CONTROL),
        "jump_to_top" => ("g", KeyModifiers::empty()),
        "jump_to_bottom" => ("G", KeyModifiers::SHIFT),
        "jump_to_row_start" => ("0", KeyModifiers::empty()),
        "jump_to_row_end" => ("$", KeyModifiers::SHIFT),
        // VIM-style actions
        "quit" => ("q", KeyModifiers::empty()),
        "copy_cell" => ("y", KeyModifiers::empty()),
        "copy_row" => ("Y", KeyModifiers::SHIFT),
        // Keep standard bindings for non-VIM actions
        _ => return get_default_keybinding(action),
    };

    parse_key_string(binding.0).map(|(code, _)| (code, binding.1))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_key_string() {
        assert_eq!(
            parse_key_string("q"),
            Some((KeyCode::Char('q'), KeyModifiers::empty()))
        );
        assert_eq!(
            parse_key_string("Ctrl+g"),
            Some((KeyCode::Char('g'), KeyModifiers::CONTROL))
        );
        assert_eq!(
            parse_key_string("Enter"),
            Some((KeyCode::Enter, KeyModifiers::empty()))
        );
        assert_eq!(
            parse_key_string("Shift+Tab"),
            Some((KeyCode::Tab, KeyModifiers::SHIFT))
        );
    }

    #[test]
    fn test_default_config() {
        let config = Config::default();
        assert_eq!(config.theme.default, "Default");
        assert_eq!(config.ui.max_rows, 50);
        assert_eq!(config.keybindings.profile, "default");
    }

    #[test]
    fn test_vim_keybindings() {
        assert_eq!(
            get_vim_keybinding("up"),
            Some((KeyCode::Char('k'), KeyModifiers::empty()))
        );
        assert_eq!(
            get_vim_keybinding("down"),
            Some((KeyCode::Char('j'), KeyModifiers::empty()))
        );
    }
}
