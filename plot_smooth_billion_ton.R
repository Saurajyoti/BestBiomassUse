rm(list = ls())

library(tidyverse)
library(readr)

fpath <- 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data'
ppath <- 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\figs'
fname <- 'Billion Ton Results_Best_Use.csv'

d <- read_csv(paste0(fpath, '\\', fname)) %>%
  filter(Scenario %in% c("Basecase, all energy crops")) %>%
  select(c(Year, Scenario, `Biomass Price`, Feedstock, Production, `Production Unit`, `Yield Unit`, Yield)) %>%
  filter(!is.na(Yield)) %>%
  mutate(`Biomass Price` = as.factor(`Biomass Price`),
         Feedstock_unit = paste0(Feedstock, ' (', `Production Unit`, ')'))

p <- ggplot(d) +
  geom_smooth(aes(Year, Production, color = `Biomass Price`), se = FALSE) +
  facet_wrap(~Feedstock_unit, scales = "free") +
  labs(x = "", y = "Production", 
       title = "Billion-Ton projected energy feedstocks-biomass availability",
       color = "USD per unit prod.") + 
  theme_classic() +
  theme(plot.title = element_text(hjust = 0.5),
        text = element_text(size = 16))

ggsave(paste0(ppath, '\\', 'Billion_Ton_feedstocks_smooth.jpg'), p, width = 16, height = 9, unit = "in", dpi = 300)


d_misc <- d %>%
  filter(Feedstock %in% c("Miscanthus"))

p <- ggplot(d) +
  geom_smooth(aes(Year, Production, color = `Biomass Price`), se = FALSE) +
  labs(x = "", y = "Production", 
       title = "Billion-Ton projected Miscanthus availability",
       color = "USD per unit prod.") + 
  theme_classic() +
  theme(plot.title = element_text(hjust = 0.5),
        text = element_text(size = 16))

ggsave(paste0(ppath, '\\', 'Billion_Ton_Miscanthus_smooth.jpg'), p, width = 16, height = 9, unit = "in", dpi = 300)
