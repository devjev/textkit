use crate::{PAT_HB_CPX, PAT_HB_MLE, PAT_HB_MLS, PAT_HB_SMP};
use regex::Regex;

#[derive(Debug)]
pub(crate) enum AccumulationMode {
    Uncertain,
    Simple,
    Multiline { depth: usize },
}

#[derive(Debug)]
pub(crate) enum TemplateAccumulator {
    Idle,
    Accumulating {
        acc: String,
        acc_state: AccumulationMode,
    },
    Done(String),
}

impl TemplateAccumulator {
    pub fn accumulate(&mut self, text: &str) {
        let start_uncertain = Regex::new(r"\{").unwrap(); // let start_uncertain = Regex::new(r"\{$").unwrap();
        let start_simple = Regex::new(r"\{\{").unwrap();
        let start_multi = Regex::new(r"\{\{#").unwrap();
        let end_multi = Regex::new(r"\{\{/").unwrap();
        let end_simple = Regex::new(r"\}\}").unwrap();

        let simple_ph_well_formed = Regex::new(PAT_HB_SMP).unwrap();
        let complex_ph_well_formed = Regex::new(PAT_HB_CPX).unwrap();
        let multi_ph_start_well_formed = Regex::new(PAT_HB_MLS).unwrap();
        let multi_ph_end_well_formed = Regex::new(PAT_HB_MLE).unwrap();

        *self = match self {
            Self::Idle => {
                // Here we need to decide which case we are handling:
                // 1. Is it multiline/directive placeholder (well-formed, start or end)?
                // 2. Is it simple placeholder (well-formed)?
                // 3. Is it a malformed placeholder?
                // 4. Is it nothing of the sort?
                if multi_ph_start_well_formed.is_match(text) {
                    Self::Accumulating {
                        acc: text.into(),
                        acc_state: AccumulationMode::Multiline { depth: 0 },
                    }
                } else if multi_ph_end_well_formed.is_match(text)
                    || simple_ph_well_formed.is_match(text)
                    || complex_ph_well_formed.is_match(text)
                {
                    Self::Done(text.into())
                } else if start_uncertain.is_match(text) {
                    Self::Accumulating {
                        acc: text.into(),
                        acc_state: AccumulationMode::Uncertain,
                    }
                } else {
                    Self::Done(text.into())
                }
            }
            Self::Accumulating {
                acc,
                acc_state: AccumulationMode::Multiline { .. },
            } => {
                let mut new_str = acc.clone();
                new_str.push_str(text);
                let opening_brackets = start_multi.find_iter(&new_str).count();
                let closing_brackets = end_multi.find_iter(&new_str).count();
                let still_open = opening_brackets - closing_brackets;

                if still_open > 0 {
                    Self::Accumulating {
                        acc: new_str,
                        acc_state: AccumulationMode::Multiline { depth: still_open },
                    }
                } else if still_open == 0 {
                    Self::Done(new_str)
                } else {
                    // TODO error here?
                    Self::Accumulating {
                        acc: new_str,
                        acc_state: AccumulationMode::Multiline { depth: still_open },
                    }
                }
            }
            Self::Accumulating {
                acc,
                acc_state: AccumulationMode::Simple,
            } => {
                let mut new_str = acc.clone();
                new_str.push_str(text);
                // Terminating case should be working here!!!
                if end_simple.is_match(&new_str)
                    && is_simple_placeholders_completely_closed(&new_str)
                {
                    Self::Done(new_str)
                } else {
                    Self::Accumulating {
                        acc: new_str,
                        acc_state: AccumulationMode::Simple,
                    }
                }
            }
            Self::Accumulating {
                acc,
                acc_state: AccumulationMode::Uncertain,
            } => {
                // 1. Check if it's an incomplete multiline placeholder
                // 2. Check if it's a simple placeholder
                // 3. Check if it's a finished multiline placeholder
                // 4. Check if it's a finished simple placeholder
                // 5. Nothing
                let mut new_str = acc.clone();
                new_str.push_str(text);

                // Multi-line, complete
                if multi_ph_start_well_formed.is_match(&new_str)
                    && multi_ph_end_well_formed.is_match(&new_str)
                {
                    Self::Done(new_str)
                } else if start_multi.is_match(&new_str) {
                    Self::Accumulating {
                        acc: new_str,
                        acc_state: AccumulationMode::Multiline { depth: 0 },
                    }
                } else if simple_ph_well_formed.is_match(&new_str) {
                    Self::Done(new_str)
                } else if start_simple.is_match(&new_str) {
                    Self::Accumulating {
                        acc: new_str,
                        acc_state: AccumulationMode::Simple,
                    }
                } else {
                    Self::Done(new_str.into())
                }
            }
            Self::Done(x) => Self::Done(x.clone()),
        }
    }

    pub fn reset(&mut self) {
        *self = Self::Idle;
    }
}

fn is_simple_placeholders_completely_closed(text: &str) -> bool {
    let no_of_openning_brackets = text.matches("{{").count();
    let no_of_closing_brackets = text.matches("}}").count();
    (no_of_openning_brackets - no_of_closing_brackets) == 0
}
