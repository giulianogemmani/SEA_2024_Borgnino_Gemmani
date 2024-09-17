import os
import matlab
import matlab.engine

print("Current Working Directory:", os.getcwd())

# Start MATLAB engine
eng = matlab.engine.start_matlab()
eng.addpath("C:\\Users\\borgn\\OneDrive\\Desktop\\progetto_software\\SEA_2024_Borgnino_Gemmani")
eng.matlab_main('example_simulink_scheme2023b')