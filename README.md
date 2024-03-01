# pptx-scripts
Some PPTX scripts written in Python for personal use.

## Usage
You need [python-pptx](https://pypi.org/project/python-pptx/) to run these scripts. If you haven't installed it already, install with `pip install python-pptx`

### Export notes

`python export_pptx_notes.py <input_presentation.pptx>`

This script create a new folder named as `<input_presentation>_notes` in the folder which the input pptx file is located and output notes from each slide. 

The folder structure ***after*** running the script will be:

```
├── input_presentation.pptx
├── input_presentation_notes
│   ├── Slide 1.txt
│   ├── Slide 2.txt
│   └── ...
```

### Insert autoplay audio into PPTX
`python insert_audio.py <input_presentation.pptx>`

This script will scan all the mp3 files named as `Slide X.mp3` (X being the index number) from the `audio` folder which is located at the same folder as the input pptx file, insert them into corresponding slides, make them autoplay with the slides and save as a new file named as `<input_presentation>_with_audio.pptx`. 

The folder structure ***before*** running the script should be:

```
├── input_presentation.pptx
├── audio
│   ├── Slide 1.mp3
│   ├── Slide 2.mp3
│   └── ...
```
