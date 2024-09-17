import os
import matlab
import matlab.engine

print("Current Working Directory:", os.getcwd())

# Start MATLAB engine
eng = matlab.engine.start_matlab()
eng.addpath("<TARGETDIR>")
eng.matlab_main('<TARGET>')