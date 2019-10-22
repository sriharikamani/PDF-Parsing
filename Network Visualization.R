###########################################################################################################################################
# Program Name  : Network
# Purpose       : R - Program to plot 2-mode and 1-mode network diagrams and derive centrality measurements for 3GPP S1 meeting discussions  
############################################################################################################################################
# TO remove the objects stored in workspace
rm(list=ls(all=T)) 
cat("\014")

library(igraph)
library(xlsx)
library(readxl)
library(sqldf)
library(openxlsx)
#
############################################################################################################################################
# Set the Path 
############################################################################################################################################
#
path <- paste(("C:/Users/SRIHARI/Desktop/Network Analytics/S1"))
setwd(path)
#
####################
# Load the PDF File 
####################
##
data    <- read_excel("S1_80.xlsx", sheet = 1)
###########################
#Remove first 4 characters
###########################
#
data$ID <- sub('......', '', data$ID) 
head(data)
df      <- data.frame(data$ID,data$Source)
#
###################
# Check the data   
###################
#
net <- graph.data.frame(df,directed = FALSE)
V(net)  # Vertices
E(net)  # Edges
V(net)$label = V(net)$name
V(net)$degree = degree(net)
V(net)$label
V(net)$degree
#
###############################
# Create Network Visualization
###############################
#
set.seed(222)
net2 = simplify(net,remove.multiple = TRUE,remove.loops = FALSE)

plot(net2,
     vertex.color = "orange",
     vertex.label.color='black',
     vertex.label.font=2,
     vertex.size = V(net2)$degree *0.2,
     vertex.size = degree(net2,mode = "in"),
     vertex.label.dist = 0.2,
     edge.arrow.size = 0.05,
     edge.color="blue",
     vertex.label.cex = 0.4,
     layout = layout.fruchterman.reingold,niter=500)

title(main = c(paste("Visualizing 2-mode S1-70 meeting data"),"\n:",paste("Relation Connect between 3GPP Members and TDoc")),
      cex.main = 2,   font.main= 4, col.main= "blue",
      cex.sub = 0.70, font.sub = 3, col.sub = "red")

#
################################
# Two Mode Network Visualization
################################
# 
affiliation_data             <- read.csv(file = 'S1_80_2Mode.csv', header = T,row.names = 1) # OR row.names = TRUE
affiliation_matrix           <- as.matrix(affiliation_data) 
rownames(affiliation_matrix) <- sub('......', '', rownames(affiliation_matrix)) 
two_mode_network             <- graph.incidence(affiliation_matrix)
set.seed(222)
two_mode_network = simplify(two_mode_network,remove.multiple = TRUE,remove.loops = TRUE)

plot(two_mode_network,
     vertex.color = "orange",
     vertex.label.color='black',
     vertex.label.font=2,
     vertex.size = V(net2)$degree *0.2,
     vertex.size = degree(net2,mode = "in"),
     vertex.label.dist = 0.2,
     edge.arrow.size = 0.05,
     edge.color="blue",
     vertex.label.cex = 0.4,
     layout = layout.fruchterman.reingold,niter=500)

title(main = c(paste("Visualizing 2-mode S1-70 meeting data"),"\n:",paste("Relation Connect between 3GPP Members and TDoc")),
      cex.main = 2,   font.main= 4, col.main= "blue",
      cex.sub = 0.70, font.sub = 3, col.sub = "red")

#
################################
# One Mode Network Visualization
################################
# 
one_mode_network   <- bipartite.projection(two_mode_network) # PURPOSE Convert a 2-mode dataset into a 1-mode adjacency matrix
#
##################################################################
#Get the connections between groups
##################################################################
#
get.adjacency(one_mode_network$proj1,sparse = FALSE,attr="weight")
get.adjacency(one_mode_network$proj2,sparse = FALSE,attr="weight")

########################################
# Delete the vertices who have no degree
########################################
one_mode_network$proj2 = delete.vertices(one_mode_network$proj2, V(one_mode_network$proj2)[degree(one_mode_network$proj2)==0])

#### Plot the netwrok

plot(one_mode_network$proj2,
     edge.label=E(one_mode_network$proj2)$weight,
     #edge.width=E(one_mode_network$proj2)$weight, # Line thickness based on the weight
     edge.curved = TRUE,
     #vertex.shape="circle",
     vertex.shape="none",
     vertex.color = "skyblue",
     vertex.label.color=rgb(0,0,.2,.6), 
     vertex.label.dist = 1.5,  
     vertex.label.degree=0,
     edge.arrow.size = 0.05,
     vertex.label.cex = .7,
     edge.curved=0.2,
     vertex.size = V(net)$degree[2] *0.5,
     layout = layout.fruchterman.reingold,niter=500)
     #layout = layout_nicely)

title(main = c(paste("Visualizing 1-mode S1-70 meeting data"),"\n",paste("Common TDocs relation between 3GPP Members")),
      cex.main = 2,   font.main= 4, col.main= "blue",
      cex.sub = 0.70, font.sub = 3, col.sub = "red")
#
################################################################
# Members common interest in TDoc)  weighted connections
################################################################
#
df <- get.data.frame(one_mode_network$proj2)
df <- df[order(df$weight, decreasing = TRUE), ]
colnames(df)   = c("From_Member","To_Member","Common_TDocs" )
#
#######################################
# Get total TDoc ties for each Member 
#######################################
#
qry_stmt    <- paste("SELECT From_Member,sum(Common_TDocs) from df group by From_Member ORDER BY sum(Common_TDocs)",sep="")
From_Member <- sqldf(qry_stmt)
From_Member <- From_Member[order(From_Member$`sum(Common_TDocs)`, decreasing = TRUE), ]

qry_stmt    <- paste("SELECT To_Member,sum(Common_TDocs) from df group by To_Member ORDER BY sum(Common_TDocs)",sep="")
To_Member   <- sqldf(qry_stmt)
To_Member   <- To_Member[order(To_Member$`sum(Common_TDocs)`, decreasing = TRUE), ]

rm(Consolidated)
colnames(From_Member) = c('Member','TDocs')
colnames(To_Member) = c('Member','TDocs')
Consolidated <- rbind(From_Member,To_Member)


qry_stmt     <- paste("SELECT Member,sum(TDocs) from Consolidated group by Member ORDER BY sum(TDocs)",sep="")
Consolidated <- sqldf(qry_stmt)
Consolidated <- Consolidated[order(Consolidated$`sum(TDocs)`, decreasing = TRUE), ]
colnames(Consolidated) = c('Member','TDocs')

OUT <- createWorkbook()
addWorksheet(OUT, "IndividualMemberTies")
addWorksheet(OUT, "ConsolidatedMemberTies")
writeData(OUT, sheet = "IndividualMemberTies", x = df)
writeData(OUT, sheet = "ConsolidatedMemberTies", x = Consolidated)
saveWorkbook(OUT, "TieStrength.xlsx",overwrite = TRUE)
system("taskkill /IM Excel.exe")
#
#################################################################################
# ##################### Centrality Measures #####################################
#################################################################################
# 
####################################
# Degree Centrality (number of ties)
####################################
# 
net2 <-graph.data.frame(data, directed=F)
deg <- igraph::degree(net2)
deg <- data.frame(sort(deg,decreasing = TRUE))
head(deg,10)                             #Top 10 members having maximum connects 
V(net)$name[degree(net)==max(degree(net))]  # Member having high connects
#
#############################################
# Closeness Centrality(importance/centrality)
#############################################
# 
net2 <-graph.data.frame(data, directed=T)
nearness  = as.data.frame(closeness(net2,mode = 'in'))
max(nearness)
nearness <- igraph::closeness(net2,mode = 'in')
nearness <- data.frame(sort(nearness,decreasing = TRUE))
head(nearness,3)
V(net)$name[closeness(net2,mode = 'in')==max(closeness(net2,mode = 'in'))]  # Member having high connects
#
################################################################################################################
# Betweenness Centrality (influenctial/How connective) (centrality based on a broker position connecting others)
################################################################################################################
# 
net2 <-graph.data.frame(data, directed=F)
#between = betweenness(net2)
between <- igraph::betweenness(net2)
between <- data.frame(sort(between,decreasing = TRUE))
head(between,3)
V(net)$name[betweenness(net2)==max(betweenness(net2))]  # Member having high connects
#
################################################################################################################
# Eigenvector centrality (influenctial)
################################################################################################################
# 
evc <- igraph::eigen_centrality(net2)
evc <- (sort(evc$vector,decreasing = TRUE))
head(evc,3)
#
#######################################################################################################################################
################################ E N D   O F   T H E   P R O G R A M  #################################################
#######################################################################################################################################
