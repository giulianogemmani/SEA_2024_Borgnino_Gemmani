import os
import matlab
import matlab.engine

print("Current Working Directory:", os.getcwd())

# Start MATLAB engine
eng = matlab.engine.start_matlab()
eng.addpath("C:\\Data_transfer")
eng.matlab_main('example_simulink_scheme2023b')