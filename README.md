# Auto Powerpoint Report Generator

A proof of concept that uses Python code to autommatically generate a Powerpoint report based on data provided.  Note that this program is fixed to create defined charts and Powerpoint slides based on hard written code - running this on a dataset and files different than ones provided in this repo will **not** work.

Feel free to look into the code for further comprehension/learning.  Clone yourself a copy and revise it to your needs.  If you have suggestions on improvement, feel free to send a pull request.

### To run this program:

1. Create virtual environment by entering into command terminal `python -m venv {name of your folder containing virtual environment}`.
2. Activate environment with `.\{name of environment folder}\Scripts\activate`.
3. Clone repository to your local by downloading folder to environment or entering command in terminal `git clone https://github.com/jedi3nigma/Auto-PPTX-Report-Generator.git`.
4. Change directory to downloaded repo and run `pip install -r requirements.txt` to install package dependencies.
5. Run `python main.py`.  A Powerpoint file should be generated along with relevant images in the "img" folder.

### Repo contents/outputs:
- _Sample_:  contains generated files from program as reference
- _img_: contains default images (logos) for Powerpoint report
- _ReportGen_: module with python files for generating report
- _data_: preset data used for report
- _main.py_: main report program
- _design_template.pptx_: required for master report template
