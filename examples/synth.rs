#[macro_use]
extern crate vst;
extern crate vst_gui;

use std::f32::consts::PI;
use std::sync::{Arc, Mutex};

use vst::buffer::AudioBuffer;
use vst::editor::Editor;
use vst::plugin::{Category, Plugin, Info};

const HTML: &'static str = r#"
    <!doctype html>
    <head>
        <meta charset="utf-8">
        <meta http-equiv="x-ua-compatible" content="ie=edge">
        <title></title>
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <style type="text/css">
            body {
                font-family: sans-serif;
                padding-top: 10%;
                text-align: center;
            }
        </style>
    </head>
        <body>
            <label for="waveformRange">Sine — Square</label>
            <br/>
            <input id="waveformRange" type="range" min="0" max="1.0" value="0" step="0.01"/>
            <br/>
            <label for="frequencyRange">Frequency</label>
            <br/>
            <input id="frequencyRange" type="range" min="55" max="880" value="440" step="any"/>
        </body>

        <script>
            var waveformRange = document.getElementById("waveformRange");
            var frequencyRange = document.getElementById("frequencyRange");

            waveformRange.addEventListener("change", function(event) {
                external.invoke("setWaveform " + event.target.value);
            });
            frequencyRange.addEventListener("change", function(event) {
                external.invoke("setFrequency " + event.target.value);
            });
        </script>
    </html>
"#;

struct Oscillator {
    pub frequency: f32,
    pub waveform: f32,
    pub phase: f32,
    pub amplitude: f32,
}

fn create_javascript_callback(
    oscillator: Arc<Mutex<Oscillator>>) -> vst_gui::JavascriptCallback
{
    Box::new(move |message: String| {
        let mut tokens = message.split_whitespace();

        let command = tokens.next().unwrap_or("");
        let argument = tokens.next().unwrap_or("").parse::<f32>();

        if argument.is_ok() {
            match command {
                "setWaveform" => {
                    oscillator.lock().unwrap().waveform = argument.unwrap();
                },
                "setFrequency" => {
                    oscillator.lock().unwrap().frequency = argument.unwrap();
                },
                _ => {}
            }
        }

        String::new()
    })
}

struct ExampleSynth {
    sample_rate: f32,
    // We access this object both from a UI thread and from an audio processing
    // thread.
    oscillator: Arc<Mutex<Oscillator>>,
    gui: vst_gui::PluginGui,
}

impl Default for ExampleSynth {
    fn default() -> ExampleSynth {
        let oscillator = Arc::new(Mutex::new(
            Oscillator {
                frequency: 440.0,
                waveform: 0.0,
                phase: 0.0,
                amplitude: 0.1,
            }
        ));

        ExampleSynth {
            sample_rate: 44100.0,
            oscillator: oscillator.clone(),
            gui: vst_gui::new_plugin_gui(
                String::from(HTML),
                create_javascript_callback(oscillator.clone())),
        }
    }
}

impl Plugin for ExampleSynth {
    fn get_info(&self) -> Info {
        Info {
            name: "Example Synth".to_string(),
            vendor: "rust-vst-gui".to_string(),
            unique_id: 9614,
            category: Category::Synth,
            inputs: 2,
            outputs: 2,
            parameters: 0,
            initial_delay: 0,
            f64_precision: false,
            ..Info::default()
        }
    }

    fn set_sample_rate(&mut self, sample_rate: f32) {
        self.sample_rate = sample_rate as f32;
    }

    fn process(&mut self, buffer: &mut AudioBuffer<f32>) {
        let mut oscillator = self.oscillator.lock().unwrap();

        let actual_phase = oscillator.phase;
        let actual_frequency = oscillator.frequency;

        let phase = |sample_index: usize| {
            actual_phase + 2.0 * PI * actual_frequency *
                (sample_index as f32) / self.sample_rate
        };

        for (_, output) in buffer.zip() {
            for (index, sample) in output.iter_mut().enumerate() {
                let sine_wave = phase(index).sin();
                let square_wave = phase(index).cos().signum();

                *sample = oscillator.amplitude *  (
                    sine_wave * (1.0 - oscillator.waveform) +
                    square_wave * oscillator.waveform);
            }
        }

        oscillator.phase = phase(buffer.samples()) % (2.0 * PI);
    }

    fn get_editor(&mut self) -> Option<&mut Editor> {
        Some(&mut self.gui)
    }
}

plugin_main!(ExampleSynth);
