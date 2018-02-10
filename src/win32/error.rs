use std::error::Error;
use std::fmt;

#[derive(Debug)]
pub struct RuntimeError {
    description: String,
}

impl RuntimeError {
    pub fn new(description: &str) -> RuntimeError {
        RuntimeError {description: String::from(description)}
    }
}

impl Error for RuntimeError {
    fn description(&self) -> &str {
        &self.description
    }

    fn cause(&self) -> Option<&Error> {
        None
    }
}

impl fmt::Display for RuntimeError {
    fn fmt(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        write!(formatter, "{}", self.description)
    }
}
