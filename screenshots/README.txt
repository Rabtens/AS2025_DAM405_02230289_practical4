Save your screenshots in this folder with EXACTLY these names (png or jpg), then run:
    python code/build_report.py
They are embedded into the Word report automatically.

fig01_install.png       pip install "flaml[automl]" scikit-learn  -> 'Successfully installed ...'
fig02_data_split.png    'TASK 1' block (shapes / class counts / split)
fig03_baseline.png      'TASK 2' block ('Baseline accuracy: ...')
fig04_automl_run.png    'TASK 3' block: FLAML log lines + 'AutoML search finished in ... s'
fig05_best_model.png    'TASK 4' block: best estimator, best config, best CV score
fig06_comparison.png    'TASK 5' block: AutoML vs baseline test accuracy
fig07_saved_model.png   'TASK 6' block + `ls -lh outputs/` showing automl_model.joblib
fig08_extras.png        (optional) 'OPTIONAL EXPERIMENTS' block from --extras
