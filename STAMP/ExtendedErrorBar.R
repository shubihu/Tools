library(ggplot2)
library(reshape2)
library(ggpubr)

args <- commandArgs(T)
file <- as.character(args[1])
# file <- 'E:/Desktop/t-test.csv'
df <- read.csv(file, check.names = F)

df <- df[df['Pvalue'] < 0.05,]
df <- na.omit(df)
df['Mean'] <- (df[,2] + df[,3]) / 2
df <- df[order(df['Mean']),]

g1 <- gsub('_MeanRel\\(%\\)', '', colnames(df)[2])
g2 <- gsub('_MeanRel\\(%\\)', '', colnames(df)[3])
group <- c()
for (i in 1:nrow(df)){
        g1_mean <- df[i,2]
        g2_mean <- df[i,3]
        group[i] <- ifelse(g1_mean > g2_mean, g1, g2)
}
df['Group'] <- group

newdf <- df[,c(1,2,3)]
colnames(newdf) <- c('Genus', g1, g2)
newdf['Genus'] <- factor(newdf$Genus, levels = newdf$Genus)
newdf <- melt(newdf)

p1 <- ggplot(newdf, aes(x=Genus, y=value, fill=variable)) + 
        geom_bar(stat ="summary",fun.y = mean,width = 0.3, color = "black", position = position_dodge()) +
        theme_bw() +
        theme(panel.grid.minor = element_blank(),
              panel.grid.major = element_blank(),
              legend.title=element_blank(),
              legend.position='top') +
        labs(x = "", y = 'Mean Proportion') +
        coord_flip()

colnames(df)[c(5,6,7)] <- c('dp','lowerCI', 'upperCI')
df['pvalue'] <- seq(0,1, 1 / nrow(df))[-1]
p2 <- ggplot(df, aes(x=pvalue, y=dp, color=Group)) + 
        geom_point() +
        scale_shape_manual(values = 19) +
        geom_errorbar(aes(ymin=lowerCI, ymax=upperCI), width=0.025,
                      position=position_dodge(0.05)) +
        theme_bw() +
        theme(panel.grid.minor = element_blank(),
              panel.grid.major = element_blank(),
              legend.title = element_blank()) +
        labs(x = 'pvalue', y = 'Difference in mean proportion') +
        scale_x_continuous(breaks = seq(0, 1, 1 / nrow(df))[-1], labels = format(df$Pvalue, digits = 2), position = "top") +
        geom_hline(aes(yintercept=0), colour="#E69F00", linetype="dashed") +
        theme(legend.position="top") +
        
        coord_flip()

ggarrange(p1,p2,widths=c(3/5, 2/5),ncol = 2, nrow = 1)
if (grepl('FunctionPrediction',file)){
    ggsave(gsub('csv', 'png', file), width = 30, height = 15, units = 'cm', dpi = 600)
}else{
    ggsave(gsub('csv', 'png', file), width = 20, height = 15, units = 'cm', dpi = 600)
}

ggsave(gsub('csv', 'pdf', file), width = 14, height = 10)

