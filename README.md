# Near Field Scanner VBA Macros for CST Studio
This repository gathers a set of VBA macros for the common microwave simulation software CST Studio.
Using these macros you can create a defined set of electrical field probes on a plane/cylinder/sphere and export the resulting simulated near-field measurement data.

The macros are useable but not thoroughly tested. The current version works for CST Studio Version 2017, but the code can easily be adapted.

## Importing the Macros
To use the macros, they have to be imported.
Under *Macros/Import Macros* you can select the files to import. 
Be aware that the file ending .mcr is used.

## Editing the Macros
Through the build-in macro editor the macros can be executed but also adapted.
Check the CST documentation for a VBA reference.

## Create a Planar Scanner
To create a set of probes on a plane use the *CreatePlanarScanner* macro.

![Alt text](https://raw.githubusercontent.com/hbartle/CSTStudio_NFScanner/master/images/planar_create.png?raw=true "Create Planar Scanner")

The plane can be aligned with either x-,y- or z-axis of the model WCS.
Parameters are:

* Number of probes in first direction
* Number of probes in second direction
* Spacing of probes in first direction
* Spacing of probes in second direction
* Distance of plane to origin

All distances have to be provided in mm. 
At the same time make sure that the CST unit for distance is set to mm as well.

After creation of the probes, you will see something like this:

![Alt text](https://raw.githubusercontent.com/hbartle/CSTStudio_NFScanner/master/images/patch_antenna_w_scanner.png?raw=true "Created Probes")

Before simulation, make sure to correctly build a mesh so that the probes are included (This is usually done automatically but can lead to problems in some templates).


## Create a Cylindrical Scanner

## Create a Spherical Scanner

## Retrieve the Results
Since the probed E-field components are unhandy to work with, the *getProbeData* macro was written.
It combines the E-field values of all probes with their position at a specific frequency and writes it into a single file that can easily be post-processed with mMATLAB.
For this to work, at least one Far-Field Monitor has to be set up with a frequency that lies inside the simulation frequency boundaries.
The macro will create a dedicated file for each Far-Field Monitor that has been created.
This enables easy comparison of Near- and Far-Field "measurements".


