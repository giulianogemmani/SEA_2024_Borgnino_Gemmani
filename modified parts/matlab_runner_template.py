import os
import matlab
import matlab.engine

print("Current Working Directory:", os.getcwd())

# Start MATLAB engine
eng = matlab.engine.start_matlab()
eng.addpath(<TARGETDIR>)
eng.progetto_software_main_240911('<TARGET>')