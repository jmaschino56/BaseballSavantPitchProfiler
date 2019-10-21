# Baseball Savant Pitch Profiler (Renamed from ArsenalReport)
Provides a report that breaks down each pitch thrown by a pitcher. It uses python libraries: MatPlotLib, Pandas, PyBaseball, Math, and Python-Docx to deliever reports.
Inputs: First Name, Last Name, Date 1, Date 2
Outputs:

•	Pitch Type - FF: Four-Seam Fastball, FT: Two-Seam Fastball, SI: Sinkker, FC: Cutter, FS: Splitter, SL: Slider, CU: Curveball, KC: Knuckle-curve, CH: Changeup, FO: Forkball, SC: Screwball, KN: Knuckleball, EP: Eephus

•	% Thrown - Frequency % of Pitch

•	Velocity - recorded in miles per hour at release.

•	Spin Rate - recorded in revolutions per minute at release.

•	Horizontal Break - horizontal movement, in inches, of the pitch between the release point and home plate, as compared to a theoretical pitch thrown at the same speed with no spin-induced movement. Calculated by Alan Nathan's Statcast conversions and then flipped upon the X axis as to give a pitcher's view.

•	Vertical Break - vertical movement, in inches, of the pitch between the release point and home plate, as compared to a theoretical pitch thrown at the same speed with no spin-induced movement. Calculated by Alan Nathan's Statcast conversions.

Horizontal and Vertical Movement explained well by Simple Sabermetrics: https://www.youtube.com/watch?v=ejCb-2wyAts&list=PLmtSuNbgQJKAMV6If2XRrC-HjqVO4Spn-&index=3

•	Tilt - Spin axis converted into clock time. As a rule of thumb, the ball will break in the direction of the number on the clock face. For Example: 6:00 is perfect top spin (classic “12 - 6” curveball), causing the ball to break down. 12:00 is perfect back spin (Four seam fastball, with no left-right movement), causing the ball to break upward relative to how it would have moved due to gravity alone – cutters are around 11:00 and sinkers are around 2:00 for a RHP, while cutters are around 1:00 and sinkers around 10:00 for a LHP. 3:00 is a “Frisbee” spinning and breaking to the right, while 9:00 is a “Frisbee” spinning and breaking to the left. Generated using Alan Nathan's Statcast conversions. Definition from Trackman's Glossary. Simple Sabermetrics video explaining it here: https://www.youtube.com/watch?v=6OQV9jOagOg&list=PLmtSuNbgQJKAMV6If2XRrC-HjqVO4Spn-&index=2

•	Spin Eff. (Spin Efficiency) - the percentage of spin that is lined up with the axis of the ball. Generated using Alan Nathan's Statcast conversions. Definition provided by Driveline.

•	Wiff Rate - measures swing and miss rate compared to total swings based on pitch type. EX: (FF Swings and Misses)/(Total # of Swings and misses from FF) * 100.

•	Weak% - Weak contact % on batted balls.

•	Topped% - Topped Ball % on batted balls.

•	Under% - Under the ball % on batted balls.

•	Burner% - Burner and flare % on batted balls. Shorted from Flare/Burner% to fit in table.

•	Solid% - Solid contact % on batted balls.

•	Barrel% - Barreled ball % on batted balls.

•	HardHit% - % of batted balls with an exit velocity of 95 mph or greater.

Weak% to Barrel% can be found at https://baseballsavant.mlb.com/csv-docs under 'launch speed angle'
