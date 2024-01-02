panel.smooth <- function (x, y, ...){
  points(jitter(x), jitter(y), pch = 1, cex = 1, col = rgb(0,0,0, alpha = 0.4))
  ok <- is.finite(x) & is.finite(y)
  if (any(ok)) {
    lines(stats::lowess(x[ok], y[ok]), col = "red", ...)
    r <- cor(x, y, use = "pairwise", method = "pearson")
    
    if(!is.na(r)){
      lml <- lm(y ~ x)
      abline(lml, col = "blue", ...)
    }
  }
}

panel.cor <- function(x, y, digits = 2, ...) {
  usr <- par("usr"); on.exit(par("usr"))
  par(usr = c(0, 1, 0, 1))
  r <- cor(x, y, use = "pairwise", method = "pearson")
  txt <- format(c(r, 0.123456789), digits = digits)[1]
  text(0.5, 0.5, txt, cex = 2)
}

panel.hist <- function(x, ...) {
  usr <- par("usr"); on.exit(par("usr"))
  par(usr = c(usr[1], usr[2], 0, 1.5))
  tax <- table(x, useNA = "no")
  if(length(tax) < 11) {
    breaks <- as.numeric(names(tax))
    y <- tax/max(tax)
    interbreak <- min(diff(breaks))*(length(tax)-1)/41
    rect(breaks-interbreak, 0, breaks + interbreak, y, col = "grey")
  } else {
    h <- hist(x, plot = FALSE)
    breaks <- h$breaks; nB <- length(breaks)
    y <- h$counts; y <- y/max(y)
    rect(breaks[-nB], 0, breaks[-1], y, col = "grey")
  }
}