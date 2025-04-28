Extract and clean embedded excel data from PowerPoint files

This code will do:
* Scans for .pptx/.pptm files.
* Opens them as zip archives.
* Extracts any files under ppt/embeddings/ that look like Excel data.
* Saves them with descriptive filenames into your chosen output folder.
* Replace al "n/a", "no data", "undefined" (case-insensitive), "-", Empty strings (""), Truly empty cells (None) with 0


** Run the App locally on the Server:
source .venv/bin/activate
streamlit run app.py
