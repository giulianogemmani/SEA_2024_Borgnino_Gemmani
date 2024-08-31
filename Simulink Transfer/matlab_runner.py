import os
import matlab
import matlab.engine

print("Current Working Directory:", os.getcwd())

# Start MATLAB engine
eng = matlab.engine.start_matlab()
eng.run('C:/Users/borgn/OneDrive/Desktop/Tesi Simulink/progetto_software_main_240827_a.m', nargout=0)