# Install OpenCV3 on Ubuntu

# Begin with a fresh install
sudo apt-get update
sudo apt-get upgrade

# Remove any previous installations of x264
sudo apt-get remove x264 libx264-dev
 
# Install dependencies 
sudo apt-get install build-essential checkinstall cmake pkg-config yasm
sudo apt-get install git gfortran
sudo apt-get install libjpeg8-dev libjasper-dev libpng12-dev
sudo apt-get install libtiff4-dev  # If running Ubuntu 14.04
sudo apt-get install libtiff5-dev  # If running Ubuntu 16.04
sudo apt-get install libavcodec-dev libavformat-dev libswscale-dev libdc1394-22-dev
sudo apt-get install libxine2-dev libv4l-dev
sudo apt-get install libgstreamer0.10-dev libgstreamer-plugins-base0.10-dev
sudo apt-get install qt5-default libgtk2.0-dev libtbb-dev
sudo apt-get install libatlas-base-dev
sudo apt-get install libfaac-dev libmp3lame-dev libtheora-dev
sudo apt-get install libvorbis-dev libxvidcore-dev
sudo apt-get install libopencore-amrnb-dev libopencore-amrwb-dev
sudo apt-get install x264 v4l-utils
 
# Optional dependencies
sudo apt-get install libprotobuf-dev protobuf-compiler
sudo apt-get install libgoogle-glog-dev libgflags-dev
sudo apt-get install libgphoto2-dev libeigen3-dev libhdf5-dev doxygen

# Install Python libraries
sudo apt-get install python-dev python-pip python3-dev python3-pip
sudo -H pip2 install -U pip numpy
sudo -H pip3 install -U pip numpy

# Install virtual environment
sudo -H pip2 install virtualenv virtualenvwrapper # Python 2
sudo -H pip3 install virtualenv virtualenvwrapper # Python 3
echo "# Virtual Environment Wrapper"  >> ~/.bashrc
echo "source /usr/local/bin/virtualenvwrapper.sh" >> ~/.bashrc
source ~/.bashrc

# Create venv for Python 2 and install libraries within
mkvirtualenv facecourse-py2 -p python2  # Python 2
workon facecourse-py2
pip install numpy scipy matplotlib scikit-image scikit-learn ipython
deactivate

# Create venv for Python 3 and install libraries within
mkvirtualenv facecourse-py3 -p python3  # Python 3
workon facecourse-py3
pip install numpy scipy matplotlib scikit-image scikit-learn ipython
deactivate

# Install Git
sudo apt install git

# Download OpenCV from Github
git clone https://github.com/opencv/opencv.git
cd opencv 
git checkout 3.3.1 
cd ..

# Download open_contrib from Github
git clone https://github.com/opencv/opencv_contrib.git
cd opencv_contrib
git checkout 3.3.1
cd ..

# Create a build directory
cd opencv
mkdir build
cd build

# Run CMake
cmake -D CMAKE_BUILD_TYPE=RELEASE \
      -D CMAKE_INSTALL_PREFIX=/usr/local \
      -D INSTALL_C_EXAMPLES=ON \
      -D INSTALL_PYTHON_EXAMPLES=ON \
      -D WITH_TBB=ON \
      -D WITH_V4L=ON \
      -D WITH_QT=ON \
      -D WITH_OPENGL=ON \
      -D OPENCV_EXTRA_MODULES_PATH=../../opencv_contrib/modules \
      -D BUILD_EXAMPLES=ON ..


# Note, if you encounter problem with install or with missing file "CMakeList.txt", 
# you may need to reinstall Cmake by first removing
sudo apt-get purge cmake
sudo apt install cmake
cmake --version

# Note, if that doesn't work either, try downloading from https://cmake.org/download/
sudo apt-get purge cmake
wget https://cmake.org/files/v3.12/cmake-3.12.0.tar.gz
tar -xvf cmake-3.12.0.tar.gz 
cd cmake-3.12.0

# Note, if troubleshooting
./configure
make

# Compile and install OpenCV with contrib modules in build
nproc                             # Find out number of CPU cores in your machine
make -j4                          # Note, substitute 4 using nproc output  
sudo make install
sudo sh -c 'echo "/usr/local/lib" >> /etc/ld.so.conf.d/opencv.conf'
sudo ldconfig

# Locate openCV's python binary (cv2.so)... usually in either site-packes or dist-packages director
find /usr/local/lib/ -type f -name "cv2*.so"
 
# Note, double check the exact path on your machine before running the following commands

# Create symlink in virtual environment (Python 2)
cd ~/.virtualenvs/facecourse-py2/lib/python2.7/site-packages
ln -s /usr/local/lib/python2.7/dist-packages/cv2.so cv2.so

# Create symlink in virtual environment (Python 3)
cd ~/.virtualenvs/facecourse-py3/lib/python3.6/site-packages
ln -s /usr/local/lib/python3.6/dist-packages/cv2.cpython-36m-x86_64-linux-gnu.so cv2.so

# Test OpenCV3 using a red eye remover application 
wget https://www.learnopencv.com/wp-content/uploads/2017/06/RedEyeRemover.zip
unzip RedEyeRemover.zip

# Compile and run C++ test code
# Note, there are backticks ( ` ) around pkg-config command not single quotes
g++ -std=c++11 removeRedEyes.cpp `pkg-config --libs --cflags opencv` -o removeRedEyes
./removeRedEyes

# Test Python test code
# Activate virtual enviroment
workon facecourse-py2
workon facecourse-py3

# Open ipython (run this command on terminal)
ipython
# Import cv2 and print version (run following commands in ipython)
import cv2
print cv2.__version__
# If OpenCV3 is installed correctly, above command should give output 3.3.1.
# Press CTRL+D to exit ipython

# Run RedEyeRemover demo
python removeRedEyes.py
deactivate

