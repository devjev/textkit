mod utils;

use wasm_bindgen::prelude::*;

use textkit_word::DocxTemplate;

// When the `wee_alloc` feature is enabled, use `wee_alloc` as the global
// allocator.
#[cfg(feature = "wee_alloc")]
#[global_allocator]
static ALLOC: wee_alloc::WeeAlloc = wee_alloc::WeeAlloc::INIT;

#[wasm_bindgen]
extern {
    fn alert(s: &str);
}

#[wasm_bindgen]
pub fn greet() {
    let bytes: Vec<u8> = vec![1, 2, 3, 4, 5];
    let _template = DocxTemplate::from_bytes(&bytes).unwrap();
}
