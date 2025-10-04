use std::collections::HashMap;
use std::fmt;

pub struct DisplayMap<'a, K, V>(pub &'a HashMap<K, V>);

impl<'a, K, V> fmt::Display for DisplayMap<'a, K, V>
where
    K: fmt::Display,
    V: fmt::Display,
{
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{{")?;
        let mut first = true;
        for (key, value) in self.0 {
            if !first {
                write!(f, ", ")?;
            }
            write!(f, "{}: {}", key, value)?;
            first = false;
        }
        write!(f, "}}")
    }
}

// Usage:
// let map = HashMap::from([("a", 1), ("b", 2)]);
// println!("{}", DisplayMap(&map));
