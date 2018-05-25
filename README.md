# VBA-for-Excel

Handy VBA functions for Excel I've written over the years. 

[Special Concat](special_concate.vb) - In a previous job after exporting survey data we would often be left with very wide Excel files. Survey questions like "Which of the following 50 makeup products do you use?" would consume 50 columns with values simply being "Yes" or "No".  This function creates a way to concate each of the "Yes" answers into a single cell pulling in what the question text was. Unclear what this is to be used for? Check out [this screenshot](https://raw.githubusercontent.com/click-here/VBA-for-Excel/master/img/specconcat.png) of how it should be implemented.

[Straight Line Func](StraightLineFunc.vb) - In a previous job I was tasked with finding a way to quickly identify in a large dataset exported from Survey Monkey how to easily weed out those who straight line their answers. That is, people who simply take matrix style questins and simply select the same value all the way down. Often times those who do this are not answering the questions honestly so we want to remove their responses from the dataset. The code below surves this purpose for any number of question groupings. The function returns the summed standard deviation accross each participants survey groupings allowing one to simply sort and remove easily those who either straight-line or nearly straight-line.
