# CenterOfGravity_Approach

One of the alternatives for finding optimal location of warehouses is to approach by implementing the center of gravity method for one or more warehouse location.

The coordinates of the sales points of a company and the weights determined on the demands in these coordinates are available in Excel.

Demand weighted center of gravity method is generated as continous model. Clustering is used in the model to find the closest warehouse location for each demand point. Clustering is used to collect demand points, and then the center of gravity approach is used to determine the best location within clusters. Each demand point is allocated the last warehouse position iteratively, and then the results are collected.

Flowchart for the algorithm can be represented as follows:

<img width="741" alt="Ekran Alıntısı" src="https://user-images.githubusercontent.com/44555928/115956215-3609db80-a504-11eb-91f6-9c13b8824a07.PNG">
