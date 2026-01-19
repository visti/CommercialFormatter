package komm_fmt

import "core:fmt"

DEFAULT_STOPWORDS :: []string {
	"rasmus sk�tt",
	"julemix",
	"dj pool",
	"Generic #",
	"FutureRecords",
	"dj deep",
	"megamix",
	"VBM_",
	"megamwx",
	"mega mix",
	"yearmix",
	"year mix",
	"live mix",
	"mixtape",
	"summermix",
	"dancemix",
	"GoMorgen Liner",
	"skøtt",
	"hit mix",
	"vi elsker",
	"pop mix",
	"dance mix",
	"mix cast",
	"mastermix",
	"retro mix",
	"top 90s",
	"summer mix",
	"megasession",
	"club mix",
	"rolf warming",
	"summerparty",
	"dancemix",
	"in the mix",
	"mashup",
	"weekendmix",
	"dj cosmo",
	"dj simonsen",
	"martin deejay",
	"dj swa",
	"dj daniel olsen",
	"PHILIZZ",
	"maassive",
	"dj kosta",
	"dj tedu",
}

DEFAULT_HEADLINES :: []string {
	"Date of Broadcasting",
	"Track starting time",
	"Track playing time",
	"Local-ID",
	"Track Title",
	"Main Artist",
	"Record Label",
	"DELETE", // no longer ingesting Gramex-ID
	"DELETE",
	"DELETE",
	"Country of Recording",
	"Year of first release",
	"ISRC-Code",
}

station :: struct {
	name:            string,
	filename:        string,
	positions:       []int,
	stopwords:       []string,
	ext:             []string,
	hasHeadlines:    bool,
	convert:         bool,
	positional:      bool,
	headlines:       []string,
	uniqueStopwords: []string,
}

Stations := []station{Bauer, Jyskfynske, Globus, Radio4, ANR, ABC}

ABC := station {
	name         = "ABC",
	positional   = false,
	ext          = {"txt"},
	hasHeadlines = false,
	headlines    = {
		"Date of Broadcasting", // 11
		"Track starting time", // 20
		"Track playing time", // 26
		"DELETE",
		"DELETE",
		"DELETE",
		"DELETE",
		"Track Title", // 98
		"Main Artist", // 149
		"DELETE",
		"DELETE", // 291
		"DELETE", // 291
		"DELETE", // 291
		"DELETE", // 291
		"DELETE", // 291
	},
}
ANR := station {
	name         = "ANR",
	positional   = true,
	positions    = {297, 291, 288, 267, 213, 200, 149, 98, 47, 26, 20, 11},
	ext          = {"den"},
	hasHeadlines = true,
	stopwords    = {" ANR ", "[SS]", "REKLAME", " PROMO ", "VEJR [", "LILLE FREDAG"},
	headlines    = {
		"Date of Broadcasting", // 11
		"Track starting time", // 20
		"Track playing time", // 26
		"DELETE",
		"Album Title", // 47
		"Track Title", // 98
		"Main Artist", // 149
		"ISRC-code", // ISRC
		"Record Label", // 200
		"Catalogue No.", // 206
		"Country of Recording", // 267
		"Year of First Release", // 288
		"DELETE", // 291
	},
}


Radio4 := station {
	name         = "Radio4",
	convert      = true,
	positional   = false,
	filename     = "",
	ext          = {"csv"},
	hasHeadlines = true,
	stopwords    = {"lydidentitet", "podcast only", "ikke udgivet", "jingle"},
	headlines    = {
		"Date of Broadcasting",
		"Track Starting Time",
		"Track Playing Time",
		"DELETE",
		"DELETE",
		"DELETE",
		"Album Title",
		"Track Title",
		"Main Artist",
		"ISRC-Code",
		"Record Label",
		"Catalogue No",
		"DELETE",
		"DELETE",
		"DELETE",
	},
}


Globus := station {
	positional   = false,
	name         = "Globus",
	filename     = "",
	ext          = {"txt"},
	hasHeadlines = false,
	headlines    = {
		"Date of Broadcasting",
		"Track starting time",
		"Track playing time",
		"Main Artist",
		"Track Title",
	},
	positions    = {15, 13},
}

Jyskfynske := station {
	positional   = true,
	name         = "Jyskfynske",
	ext          = {"txt", "den"},
	hasHeadlines = false,
	positions    = {297, 291, 288, 267, 206, 200, 149, 98, 47, 26, 20, 11},
	headlines    = {
		"Date of Broadcasting", // 11
		"Track starting time", // 20
		"Track playing time", // 26
		"DELETE",
		"Album Title", // 47
		"Track Title", // 98
		"Main Artist", // 149
		"DELETE",
		"Record Label", // 200
		"Catalogue No.", // 206
		"Country of Recording", // 267
		"Year of First Release", // 288
		"DELETE", // 291
	},
	stopwords    = {
		"vejr",
		"vejle",
		";classic ",
		";dph ",
		"classic fm",
		";rw*",
		" rv ",
		" rek ",
		" intro ",
		"gmmj",
		" acc happy",
		"sponsor",
		"viborg",
		"rv_",
		"*rv",
		"wb_",
		" sw ",
		"sw_",
		"sweeper",
		"vlr",
		"toh ",
		"sw_unknown artist",
		"skala listen",
		"skala_",
		"_skala",
		"jfm",
		"festudvalget",
		"skala fm",
		"promo",
		"nyheder",
		"reklame",
		"toh skala",
		"fest_",
		"vejrsyd",
		"vejr [",
		"listen_count",
		"dst_",
		"fa_",
		"vib_",
		"rek-",
		"vi elsker - ",
		"happy hour",
		"trackbed",
	},
}

Bauer := station {
	positional   = true,
	name         = "Bauer",
	filename     = "",
	positions    = {185, 179, 173, 168, 163, 153, 128, 78, 28, 19, 14, 6},
	hasHeadlines = false,
	ext          = {"txt"},
	stopwords    = {
		"R1 ",
		"PO Duran",
		"PODCAST",
		" R1 ",
		"TOP HOUR",
		"99999998",
		"VO *:",
		"NO News",
		"NO BED",
		"jingle",
		"SWEEPER",
		";SOFT ",
		"RADIO SOFT",
		"MR SW",
		";MR",
		" BREAKER",
		"NYHEDER",
		"PO Hitliste",
		"PROMO",
		"VEJR",
		"Bauer",
	},
	headlines    = DEFAULT_HEADLINES,
}

printstation :: proc(station: station) {
	fmt.printf(
		"Name: %v\nPositions: %v\nFilename:%v\nStopwords:\n",
		station.name,
		station.positions,
		station.filename,
	)

	for word in station.stopwords {
		fmt.printf("%v\n", word)
	}
}
