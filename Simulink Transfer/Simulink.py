'''
Created on 13 lug 2024

@author: borgn
'''
import os
import sys
#sys.path.append('C:\\Program Files\\MATLAB\\R2023b\\extern\\engines\\python')

#from subprocess import Popen
#p = Popen("activate.bat", cwd=r"C:\Users\borgn\eclipse-workspace\Scripts")
#p = Popen(["dir"])

#with Popen(["ifconfig"], stdout=PIPE) as proc:
#    log.write(proc.stdout.read())
    
import subprocess

#python_311 = subprocess.Popen("C:\\Users\\borgn\\eclipse-workspace\\Scripts\\activate.bat",  shell = True)
#command = ["C:\\Users\\borgn\\eclipse-workspace\\Scripts\\matlab_runner.py", ""]
#python_311 = subprocess.Popen(command,  shell = True)


#python_311 = subprocess.Popen("dir", shell = True, stdin=subprocess.PIPE, stdout=None, stderr=subprocess.PIPE)


#python_311 = subprocess.Popen("activate.bat", cwd = "C:\\Users\\borgn\\eclipse-workspace\\Scripts\\", stdin = subprocess.PIPE, shell = True)
#python_311.stdin.write((r"C:\\Users\\borgn\\eclipse-workspace\\Scripts\\activate.bat" + "\n").encode())
#python_311.stdin.flush()

command = []
python_311 = subprocess.Popen("activate.bat", cwd = "C:\\Users\\borgn\\eclipse-workspace\\Scripts\\", stdin = subprocess.PIPE, shell = True)

#command2 = ["python", "matlab_runner.py"]
#python_311.communicate("python matlab_runner.py".encode()) 

python_311.stdin.write('dir\n'.encode())
python_311.stdin.flush()

python_311.communicate('python matlab_runner.py\n'.encode())

#python_311 = subprocess.run(["C:\\Users\\borgn\\eclipse-workspace\\Scripts\\activate.bat"],  shell = True)


#python_311 = subprocess.Popen("python C:\\Users\\borgn\\eclipse-workspace\\Scripts\\matlab_runner.py", shell = True)

#python_311.communicate('python\n'.encode()) 

#python_311.stdin.write("dir\n".encode())
#python_311.stdin.flush()

#python_311.communicate('python matlab_runner.py\n'.encode())

#def run_job(self, jobname, level):
#    run = self.console.communicate("run job=%s level=%s yes" % (jobname, level))
#    return(run)


#python_311.stdin.write('dir\n'.encode())
#python_311.stdin.flush()
 
#result = python_311.run(["dir"], shell=True, capture_output=True, text=True)
#python_311.stdin.write('dir\n'.encode())    
#python_311.stdin.write("dir\n".encode())
#python_311.stdin.write('python matlab_runner.py\n'.encode())

#stdout, stderr = p.communicate()

#result = subprocess.run(["dir"], shell=True, capture_output=True, text=True)



#import matlab
#import matlab.engine
print("Current Working Directory Eclipse:", os.getcwd())

#from docx import Document
#from docx.shared import Inches
# Start MATLAB engine
#eng = matlab.engine.start_matlab()
# Alternatively, specify the full path if the script is not in the MATLAB path
#eng.run('C:/Users/borgn/OneDrive/Desktop/Tesi Simulink/progetto_software_main_240827_a.m', nargout=0)
# Create a new Document
#doc = Document()

# Add a Title
#doc.add_heading('MATLAB Figure Report', level=1)

# Add the figure
#doc.add_picture('figure.png', width=Inches(5))  # Adjust the width as needed

# Add a caption
#doc.add_paragraph('Figure 1: Example plot generated in MATLAB.')

# Save the document
#doc.save('C:\\Users\\borgn\\OneDrive\\Desktop\\progetto software\\report.docx')
