This project is a simple Python code that uses the location coordinates to calculate the optimal azimuth and tilt angle for PV modules.

The program calculates this by determining the sun's position for each hour throughout the entire year. Then values are obtained for Direct Normal Irradiance (DNI), Diffused Horizontal Irradiance (DHI), and Global Horizontal Irradiance (GHI) from an Excel sheet provided with the Python code.

From this information, we can estimate the total irradiance on the PV module for different azimuths and tilt angles. Then, the code sums up the total energy through all hours and shows a contour plot to give the optimum direction and tilt of the PV module.

![illustration](https://github.com/user-attachments/assets/46dd61a5-016a-408e-81e6-bdcdce6d289a)
![output](https://github.com/user-attachments/assets/dbd9e629-80dd-4d31-8e3c-9f366b960a9e)
