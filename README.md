# HP 3457A Control and Data Logging Software
This software allows you to control and log data from the HP 3457A multimeter via an AR488 Arduino GPIB Adapter. Supports Window 10, 8, and 7.

#### [Download](https://github.com/Niravk1997/HP-3457A-Software/releases)

#### [EEVblog forum post](https://www.eevblog.com/forum/testgear/hp-34401a-standalone-software/)

#### [AR488 Adapter](https://github.com/Twilight-Logic/AR488)



#### The main software window
![HP 3457A Software](https://github.com/Niravk1997/HP-3457A-Software/blob/main/Images/HP3457A_Main_Window.gif)

#### Interactive Graphing Module
![HP 3478A Graph Module](https://github.com/Niravk1997/HP-3457A-Software/blob/main/Images/HP3457A_Graph_Window.gif)

**Features:**

- Supports 7.5 Digit Resolution Mode.
- Control and Log data, save data into organized folders
- Multithreading Support:
   - All Windows open in a new thread, this ensures maximum performance. For example: Interacting with the Graph Window does not slow down other Windows.
    - All Serial communication happens on an excusive thread. This allows the software to maintain maximum sample capture speed at all times, regardless of what other data processing might be going on.
    - Users can interact with the Graph Window smoothly without any lag.
    - Data Logging functions also run on their designated thread, periodically the software will save measurement data from FIFO data structures into text files.
- Speech Synthesizer feature allows the software to voice measurements periodically and or when it meets the maximum or minimum value threshold.
- Graph Window allows users to visualize their captured data. You can get statistics for all the samples capture or for select few samples. Pan, zoom, and zoom to highlighted area. Save/copy graph as image or save graph's data into text/csv files. 
- Create math waveforms from the samples captured. Create math waveforms from math waveforms. There is no limit to how many math waveforms you can create. 
- Create Histogram from the samples captured and from math waveforms. There is no limit to how many histogram waveforms you can create.
- Measurement table allows users to collect and display the measurement data into a table.
- Capture up to 160  measurement samples in 2 seconds.

#### Create Math Waveform
![HP 3457A Math Waveform](https://github.com/Niravk1997/HP-3457A-Software/blob/main/Images/HP3457A_Math_Window.gif)

#### Histogram Waveform
![HP 3457A Histogram](https://github.com/Niravk1997/HP-3457A-Software/blob/main/Images/HP3457A_Histogram_Window.gif)

#### Measurement Table
![HP 3457A Table](https://github.com/Niravk1997/HP-3457A-Software/blob/main/Images/HP3457A_Table_Window.gif)

