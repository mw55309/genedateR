library(readxl)

# PowersRankedDeletions.xls
# LegubeSuppTableS1.xls
# GeorletteSuppTableS1.xls

ss <- "GeorletteSuppTableS1.xls"

xl <- read_excel(ss)

# iterate over columns
for (i in 1:ncol(xl)) {

	if (! typeof(xl[,i]) == "character") {
		next
	}

	
	# iterate over values
	for (j in 1:length(xl[,i])) {

		# skip if it's NULL
		if (is.null(xl[j,i])) {
			next;
		}

		# skip if it's NA
		if (is.na(xl[j,i])) {
			next;
		}

		# Excel dates when read by R
		# look something like 35764
		# which is time since "1899-12-31"

		# let's test if we can convert
		# to a date

		ad <- tryCatch(as.Date(as.numeric(xl[j,i]), "1899-12-31"),
			       error = function(e) {},
			       warning = function(w) {}

		)

		# do we have a date?
		if (inherits(ad, "Date")) {
			# we do!
			ydiff <- as.numeric(ad - as.Date("1899-12-31")) / 365

			if (is.na(ydiff)) {
				next
			}

			if (ydiff > 70 & ydiff < 118) {
				# it's a plausible date
				xldate <- format(ad, "%b-%d")
				msg <- paste("row",j,"column",i,xl[j,i],"may contain",xldate, sep=" ")
				print(msg)
			}
		}

	}

}
