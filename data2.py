import pandas as pd
import matplotlib.pyplot as plt 
import numpy as np

#Read document and save it as dataframe
df = pd.read_excel('SuperstoreDataset.xlsx')

def plotgraph(Lcat,Scat,Vals):
    #initialise graph
    fig, ax = plt.subplots()

    #set the positions of the bars on the graph
    width = 0.25
    categories = df[Lcat].unique()
    segments = df[Scat].unique()
    n_groups = len(categories)
    indices = np.arange(n_groups)

    for i, segment in enumerate(segments):
        totals = []
        for category in categories:
            #get the relevant items
            filtered_data = df[(df[Scat] == segment) & (df[Lcat] == category)]
            total_quantity = filtered_data[Vals].sum()
            totals.append(total_quantity)
        #create the bars
        rects = ax.bar(indices + i * width, totals, width, label=segment)
        #labeling
        ax.bar_label(rects, padding=3)

    ax.set_xticks(indices + width, categories)

    ax.set_xlabel(Lcat)
    ax.set_ylabel(Vals)

    ax.legend()
    #save the graph
    plt.savefig(str("./images/"+Scat+Lcat+Vals+".png"))
    #show the graph
    plt.show()

#in the form: (major category, grouped bars, values)
plotgraph("Category","Segment","Quantity")

#the consumer is the most significant part of the sales, 
#making up more than both other categories put together