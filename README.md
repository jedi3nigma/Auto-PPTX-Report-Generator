# Auto Powerpoint Report Generator

A proof of concept that uses Python code to autommatically generate a Powerpoint report based on data provided.  Note that this program is fixed to create defined charts and Powerpoint slides based on hard written code - running this on a dataset and files different than ones provided in this repo will **not** work.

Feel free to look into the code for further comprehension/learning.  Clone yourself a copy and revise it to your needs.  If you have suggestions on improvement, feel free to send a pull request.

### To run this program:

1. Clone repository to your local.
2. Create and activate virtual environment.  Run `pip install -r requirements.txt` to install package dependencies.  You may skip this if you already have required package versions of required libraries.
3. Open terminal and run `python main.py`.

### Repo contents:
- _Sample_:  contains generated files from program as reference
- _img_: contains default images (logos) for Powerpoint report
- _ReportGen_: module with python files for generating report
- _data_: preset data used for report
- _main.py_: main report program
- _design_template.pptx_: required for master report template
