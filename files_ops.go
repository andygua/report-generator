package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

//EXCEL Related global variables and constats
var xAxis = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ", "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ", "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ", "IA", "IB", "IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", "IW", "IX", "IY", "IZ", "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", "JM", "JN", "JO", "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JW", "JX", "JY", "JZ", "KA", "KB", "KC", "KD", "KE", "KF", "KG", "KH", "KI", "KJ", "KK", "KL", "KM", "KN", "KO", "KP", "KQ", "KR", "KS", "KT", "KU", "KV", "KW", "KX", "KY", "KZ", "LA", "LB", "LC", "LD", "LE", "LF", "LG", "LH", "LI", "LJ", "LK", "LL", "LM", "LN", "LO", "LP", "LQ", "LR", "LS", "LT", "LU", "LV", "LW", "LX", "LY", "LZ", "MA", "MB", "MC", "MD", "ME", "MF", "MG", "MH", "MI", "MJ", "MK", "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA", "NB", "NC", "ND", "NE", "NF", "NG", "NH", "NI", "NJ", "NK", "NL", "NM", "NN", "NO", "NP", "NQ", "NR", "NS", "NT", "NU", "NV", "NW", "NX", "NY", "NZ", "OA", "OB", "OC", "OD", "OE", "OF", "OG", "OH", "OI", "OJ", "OK", "OL", "OM", "ON", "OO", "OP", "OQ", "OR", "OS", "OT", "OU", "OV", "OW", "OX", "OY", "OZ", "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ", "RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ", "RK", "RL", "RM", "RN", "RO", "RP", "RQ", "RR", "RS", "RT", "RU", "RV", "RW", "RX", "RY", "RZ", "SA", "SB", "SC", "SD", "SE", "SF", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SP", "SQ", "SR", "SS", "ST", "SU", "SV", "SW", "SX", "SY", "SZ", "TA", "TB", "TC", "TD", "TE", "TF", "TG", "TH", "TI", "TJ", "TK", "TL", "TM", "TN", "TO", "TP", "TQ", "TR", "TS", "TT", "TU", "TV", "TW", "TX", "TY", "TZ", "UA", "UB", "UC", "UD", "UE", "UF", "UG", "UH", "UI", "UJ", "UK", "UL", "UM", "UN", "UO", "UP", "UQ", "UR", "US", "UT", "UU", "UV", "UW", "UX", "UY", "UZ", "VA", "VB", "VC", "VD", "VE", "VF", "VG", "VH", "VI", "VJ", "VK", "VL", "VM", "VN", "VO", "VP", "VQ", "VR", "VS", "VT", "VU", "VV", "VW", "VX", "VY", "VZ", "WA", "WB", "WC", "WD", "WE", "WF", "WG", "WH", "WI", "WJ", "WK", "WL", "WM", "WN", "WO", "WP", "WQ", "WR", "WS", "WT", "WU", "WV", "WW", "WX", "WY", "WZ", "XA", "XB", "XC", "XD", "XE", "XF", "XG", "XH", "XI", "XJ", "XK", "XL", "XM", "XN", "XO", "XP", "XQ", "XR", "XS", "XT", "XU", "XV", "XW", "XX", "XY", "XZ", "YA", "YB", "YC", "YD", "YE", "YF", "YG", "YH", "YI", "YJ", "YK", "YL", "YM", "YN", "YO", "YP", "YQ", "YR", "YS", "YT", "YU", "YV", "YW", "YX", "YY", "YZ", "ZA", "ZB", "ZC", "ZD", "ZE", "ZF", "ZG", "ZH", "ZI", "ZJ", "ZK", "ZL", "ZM", "ZN", "ZO", "ZP", "ZQ", "ZR", "ZS", "ZT", "ZU", "ZV", "ZW", "ZX", "ZY", "ZZ"}
var yAxis = []string{"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "100", "101", "102", "103", "104", "105", "106", "107", "108", "109", "110", "111", "112", "113", "114", "115", "116", "117", "118", "119", "120", "121", "122", "123", "124", "125", "126", "127", "128", "129", "130", "131", "132", "133", "134", "135", "136", "137", "138", "139", "140", "141", "142", "143", "144", "145", "146", "147", "148", "149", "150", "151", "152", "153", "154", "155", "156", "157", "158", "159", "160", "161", "162", "163", "164", "165", "166", "167", "168", "169", "170", "171", "172", "173", "174", "175", "176", "177", "178", "179", "180", "181", "182", "183", "184", "185", "186", "187", "188", "189", "190", "191", "192", "193", "194", "195", "196", "197", "198", "199", "200", "201", "202", "203", "204", "205", "206", "207", "208", "209", "210", "211", "212", "213", "214", "215", "216", "217", "218", "219", "220", "221", "222", "223", "224", "225", "226", "227", "228", "229", "230", "231", "232", "233", "234", "235", "236", "237", "238", "239", "240", "241", "242", "243", "244", "245", "246", "247", "248", "249", "250", "251", "252", "253", "254", "255", "256", "257", "258", "259", "260", "261", "262", "263", "264", "265", "266", "267", "268", "269", "270", "271", "272", "273", "274", "275", "276", "277", "278", "279", "280", "281", "282", "283", "284", "285", "286", "287", "288", "289", "290", "291", "292", "293", "294", "295", "296", "297", "298", "299", "300", "301", "302", "303", "304", "305", "306", "307", "308", "309", "310", "311", "312", "313", "314", "315", "316", "317", "318", "319", "320", "321", "322", "323", "324", "325", "326", "327", "328", "329", "330", "331", "332", "333", "334", "335", "336", "337", "338", "339", "340", "341", "342", "343", "344", "345", "346", "347", "348", "349", "350", "351", "352", "353", "354", "355", "356", "357", "358", "359", "360", "361", "362", "363", "364", "365", "366", "367", "368", "369", "370", "371", "372", "373", "374", "375", "376", "377", "378", "379", "380", "381", "382", "383", "384", "385", "386", "387", "388", "389", "390", "391", "392", "393", "394", "395", "396", "397", "398", "399", "400", "401", "402", "403", "404", "405", "406", "407", "408", "409", "410", "411", "412", "413", "414", "415", "416", "417", "418", "419", "420", "421", "422", "423", "424", "425", "426", "427", "428", "429", "430", "431", "432", "433", "434", "435", "436", "437", "438", "439", "440", "441", "442", "443", "444", "445", "446", "447", "448", "449", "450", "451", "452", "453", "454", "455", "456", "457", "458", "459", "460", "461", "462", "463", "464", "465", "466", "467", "468", "469", "470", "471", "472", "473", "474", "475", "476", "477", "478", "479", "480", "481", "482", "483", "484", "485", "486", "487", "488", "489", "490", "491", "492", "493", "494", "495", "496", "497", "498", "499", "500", "501", "502", "503", "504", "505", "506", "507", "508", "509", "510", "511", "512", "513", "514", "515", "516", "517", "518", "519", "520", "521", "522", "523", "524", "525", "526", "527", "528", "529", "530", "531", "532", "533", "534", "535", "536", "537", "538", "539", "540", "541", "542", "543", "544", "545", "546", "547", "548", "549", "550", "551", "552", "553", "554", "555", "556", "557", "558", "559", "560", "561", "562", "563", "564", "565", "566", "567", "568", "569", "570", "571", "572", "573", "574", "575", "576", "577", "578", "579", "580", "581", "582", "583", "584", "585", "586", "587", "588", "589", "590", "591", "592", "593", "594", "595", "596", "597", "598", "599", "600", "601", "602", "603", "604", "605", "606", "607", "608", "609", "610", "611", "612", "613", "614", "615", "616", "617", "618", "619", "620", "621", "622", "623", "624", "625", "626", "627", "628", "629", "630", "631", "632", "633", "634", "635", "636", "637", "638", "639", "640", "641", "642", "643", "644", "645", "646", "647", "648", "649", "650", "651", "652", "653", "654", "655", "656", "657", "658", "659", "660", "661", "662", "663", "664", "665", "666", "667", "668", "669", "670", "671", "672", "673", "674", "675", "676", "677", "678", "679", "680", "681", "682", "683", "684", "685", "686", "687", "688", "689", "690", "691", "692", "693", "694", "695", "696", "697", "698", "699", "700", "701", "702", "703", "704", "705", "706", "707", "708", "709", "710", "711", "712", "713", "714", "715", "716", "717", "718", "719", "720", "721", "722", "723", "724", "725", "726", "727", "728", "729", "730", "731", "732", "733", "734", "735", "736", "737", "738", "739", "740", "741", "742", "743", "744", "745", "746", "747", "748", "749", "750", "751", "752", "753", "754", "755", "756", "757", "758", "759", "760", "761", "762", "763", "764", "765", "766", "767", "768", "769", "770", "771", "772", "773", "774", "775", "776", "777", "778", "779", "780", "781", "782", "783", "784", "785", "786", "787", "788", "789", "790", "791", "792", "793", "794", "795", "796", "797", "798", "799", "800", "801", "802", "803", "804", "805", "806", "807", "808", "809", "810", "811", "812", "813", "814", "815", "816", "817", "818", "819", "820", "821", "822", "823", "824", "825", "826", "827", "828", "829", "830", "831", "832", "833", "834", "835", "836", "837", "838", "839", "840", "841", "842", "843", "844", "845", "846", "847", "848", "849", "850", "851", "852", "853", "854", "855", "856", "857", "858", "859", "860", "861", "862", "863", "864", "865", "866", "867", "868", "869", "870", "871", "872", "873", "874", "875", "876", "877", "878", "879", "880", "881", "882", "883", "884", "885", "886", "887", "888", "889", "890", "891", "892", "893", "894", "895", "896", "897", "898", "899", "900", "901", "902", "903", "904", "905", "906", "907", "908", "909", "910", "911", "912", "913", "914", "915", "916", "917", "918", "919", "920", "921", "922", "923", "924", "925", "926", "927", "928", "929", "930", "931", "932", "933", "934", "935", "936", "937", "938", "939", "940", "941", "942", "943", "944", "945", "946", "947", "948", "949", "950", "951", "952", "953", "954", "955", "956", "957", "958", "959", "960", "961", "962", "963", "964", "965", "966", "967", "968", "969", "970", "971", "972", "973", "974", "975", "976", "977", "978", "979", "980", "981", "982", "983", "984", "985", "986", "987", "988", "989", "990", "991", "992", "993", "994", "995", "996", "997", "998", "999", "1000"}
var Operations = []string{"CREATE", "CREATE-H", "CREATE-W", "READ", "READ-H", "READ-W", "UPDATE", "UPDATE-H", "UPDATE-W", "DELETE", "DELETE-H", "DELETE-W"}

const (
	CREATE = iota
	CREATE_H
	CREATE_W
	READ
	READ_H
	READ_W
	UPDATE
	UPDATE_H
	UPDATE_W
	DELETE
	DELETE_H
	DELETE_W
)

//Memtier related structs and options
type MemtierOut struct {
	Config   Configuration   `json:"configuration"`
	Runinfo  Run_information `json:"run information"`
	Allstats AllStatsStruct  `json:"ALL STATS"`
}

type SubTest struct {
	Name                    string
	memout                  []*MemtierOut
	topfilename, dufilename string
}

type RedisTest struct {
	Name string
	Sub  []SubTest
}

type Configuration struct {
	Authenticate      string  `json:"authenticate"`
	ClientStats       string  `json:"client_stats"`
	Clients           int     `json:"clients"`
	DataImport        string  `json:"data_import"`
	DataOffset        int     `json:"data_offset"`
	DataSize          int     `json:"data_size"`
	DataSizeList      string  `json:"data_size_list"`
	DataSizePattern   string  `json:"data_size_pattern"`
	DataSizeRange     string  `json:"data_size_range"`
	DataVerify        string  `json:"data_verify"`
	Debug             int     `json:"debug"`
	ExpiryRange       string  `json:"expiry_range"`
	GenerateKeys      string  `json:"generate_keys"`
	KeyMaximum        int     `json:"key_maximum"`
	KeyMedian         float64 `json:"key_median"`
	KeyMinimum        int     `json:"key_minimum"`
	KeyPattern        string  `json:"key_pattern"`
	KeyPrefix         string  `json:"key_prefix"`
	KeyStddev         float64 `json:"key_stddev"`
	MultiKeyGet       float64 `json:"multi_key_get"`
	No_expiry         string  `json:"no-expiry"`
	Num_slaves        string  `json:"num-slaves"`
	OutFile           string  `json:"out_file"`
	Pipeline          int     `json:"pipeline"`
	Port              int     `json:"port"`
	Protocol          string  `json:"protocol"`
	RandomData        string  `json:"random_data"`
	Ratio             string  `json:"ratio"`
	ReconnectInterval int     `json:"reconnect_interval"`
	Requests          int     `json:"requests"`
	RunCount          int     `json:"run_count"`
	Select_db         int     `json:"select-db"`
	Server            string  `json:"server"`
	TestTime          int     `json:"test_time"`
	Threads           int     `json:"threads"`
	Unix_socket       string  `json:"unix socket"`
	VerifyOnly        string  `json:"verify_only"`
	Wait_ratio        string  `json:"wait-ratio"`
	Wait_timeout      string  `json:"wait-timeout"`
}

type Run_information struct {
	Connections_per_thread int `json:"Connections per thread"`
	Requests_per_thread    int `json:"Requests per thread"`
	Threads                int `json:"Threads"`
}

type AllStatsStruct struct {
	Gets, Sets, Waits OpStat
	CREATE            OpStat
	CREATE_H          OpStat `json:"CREATE-H"`
	CRATE_W           OpStat `json:"CREATE-W"`
	READ              OpStat
	READ_H            OpStat `json:"READ-H"`
	READ_W            OpStat `json:"READ-W"`
	UPDATE            OpStat
	UPDATE_H          OpStat `json:"UPDATE-H"`
	UPDATE_W          OpStat `json:"UPDATE-W"`
	DELETE            OpStat
	DELETE_H          OpStat       `json:"DELETE-H"`
	DELETE_W          OpStat       `json:"DELETE-W"`
	SET               []HistoPoint `json:"SET"`
	GET               []HistoPoint `json:"GET"`
	WAIT              []HistoPoint `json:"WAIT"`
	CREATE_L          []HistoPoint `json:"CREATE-L"`
	READ_L            []HistoPoint `json:"READ-L"`
	UPDATE_L          []HistoPoint `json:"UPDATE-L"`
	DELETE_L          []HistoPoint `json:"DELETE-L"`
}

type OpStat struct {
	Hits_sec   float64 `json:"Hits/sec"`
	KB_sec     float64 `json:"KB/sec"`
	Latency    float64 `json:"Latency"`
	Misses_sec float64 `json:"Misses/sec"`
	Ops_sec    float64 `json:"Ops/sec"`
}

type HistoPoint struct {
	Msec    float64 `json:"<=msec"`
	Percent float64 `json:"percent"`
}

func read_memtier_json(fileName string) (*MemtierOut, error) {

	var mOutput MemtierOut

	json_data, err := ioutil.ReadFile(fileName)
	if err != nil {
		fmt.Printf("ERROR: occured while reading file = %v\n", err)
		return nil, err
	}

	json_data = bytes.Replace(json_data, []byte("-nan"), []byte("0.00"), -1)

	err = json.Unmarshal(json_data, &mOutput)
	if err != nil {
		fmt.Printf("ERROR: occured while unmarshalling json = %v\n", err)
		return nil, err
	}

	//fmt.Printf("json = %+v\n", mOutput)

	return &mOutput, nil

}

//
//Excel related functions starts here --------------
//
func getBlackBorderCenter(xlsx *excelize.File) int {
	s1, err := xlsx.NewStyle(`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"alignment":{"horizontal":"center", "vertical":"center"}}`)
	if err != nil {
		fmt.Printf("Error generating style() = %v\n", err)
		return -1
	}
	return s1
}

func getThroughput(o *MemtierOut, op int) float64 {
	var value float64
	switch op {
	case CREATE:
		value = o.Allstats.CREATE.Ops_sec
	case CREATE_H:
		value = o.Allstats.CREATE_H.Ops_sec
	case CREATE_W:
		value = o.Allstats.CRATE_W.Ops_sec
	case READ:
		value = o.Allstats.READ.Ops_sec
	case READ_H:
		value = o.Allstats.READ_H.Ops_sec
	case READ_W:
		value = o.Allstats.READ_W.Ops_sec
	case UPDATE:
		value = o.Allstats.UPDATE.Ops_sec
	case UPDATE_H:
		value = o.Allstats.UPDATE_H.Ops_sec
	case UPDATE_W:
		value = o.Allstats.UPDATE_W.Ops_sec
	case DELETE:
		value = o.Allstats.DELETE.Ops_sec
	case DELETE_H:
		value = o.Allstats.DELETE_H.Ops_sec
	case DELETE_W:
		value = o.Allstats.DELETE_W.Ops_sec
	}
	return value
}

func getLatency(o *MemtierOut, op int) float64 {
	var value float64
	switch op {
	case CREATE:
		value = o.Allstats.CREATE.Latency
	case CREATE_H:
		value = o.Allstats.CREATE_H.Latency
	case CREATE_W:
		value = o.Allstats.CRATE_W.Latency
	case READ:
		value = o.Allstats.READ.Latency
	case READ_H:
		value = o.Allstats.READ_H.Latency
	case READ_W:
		value = o.Allstats.READ_W.Latency
	case UPDATE:
		value = o.Allstats.UPDATE.Latency
	case UPDATE_H:
		value = o.Allstats.UPDATE_H.Latency
	case UPDATE_W:
		value = o.Allstats.UPDATE_W.Latency
	case DELETE:
		value = o.Allstats.DELETE.Latency
	case DELETE_H:
		value = o.Allstats.DELETE_H.Latency
	case DELETE_W:
		value = o.Allstats.DELETE_W.Latency
	}
	return value
}

func plotHistogram(xlsx *excelize.File, output []*MemtierOut, sheet string, cell string) error {

	var ccat, rcat, ucat, dcat string
	var cval, rval, uval, dval string
	for i, v := range output {
		ccat, rcat, ucat, dcat = "{", "{", "{", "{"
		cval, rval, uval, dval = "{", "{", "{", "{"

		for _, c := range v.Allstats.CREATE_L {
			ccat = ccat + fmt.Sprintf("%f,", c.Percent)
			cval = cval + fmt.Sprintf("%f,", c.Msec)
		}
		for _, r := range v.Allstats.READ_L {
			rcat = rcat + fmt.Sprintf("%f,", r.Percent)
			rval = rval + fmt.Sprintf("%f,", r.Msec)
		}
		for _, u := range v.Allstats.UPDATE_L {
			ucat = ucat + fmt.Sprintf("%f,", u.Percent)
			uval = uval + fmt.Sprintf("%f,", u.Msec)
		}
		for _, d := range v.Allstats.DELETE_L {
			dcat = dcat + fmt.Sprintf("%f,", d.Percent)
			dval = dval + fmt.Sprintf("%f,", d.Msec)
		}

		ccat, rcat, ucat, dcat = ccat+"0.0}", rcat+"0.0}", ucat+"0.0}", dcat+"0.0}"
		cval, rval, uval, dval = cval+"0.0}", rval+"0.0}", dval+"0.0}", dval+"0.0}"

		chartFormatStr := fmt.Sprintf(`{"type":"scatter","series":[{"name":"\"create\"","categories":"%s","values":"%s"},{"name":"\"read\"","categories":"%s","values":"%s"},{"name":"\"update\"","categories":"%s","values":"%s"},{"name":"\"delete\"","categories":"%s","values":"%s"}],"title":{"name":"Histogram-%d"}}`,
			cval, ccat, rval, rcat, uval, ucat, dval, dcat, i)

		//fmt.Printf("chart fromat string = %s\n", chartFormatStr)
		err := xlsx.AddChart(sheet, cell, chartFormatStr)
		if err != nil {
			fmt.Printf("Error generating chart err = %v\n", err)
			return err
		}
	}
	return nil
}

func plotTop(xlsx *excelize.File, topfilename string, sheet string, cellMemory string, cellCPU string) error {

	var topvalsMem []string
	var topvalCPU []string
	var skipAlternatedata bool

	//fmt.Printf("processing %s top file\n", topfilename)
	topData, err := ioutil.ReadFile(topfilename)
	if err != nil {
		fmt.Printf("Error opening the file = %s err=%v\n", topfilename, err)
		return err
	}
	topDataStr := string(topData)

	lines := strings.Split(topDataStr, "\n")
	fmt.Printf("top processing %d lines\n", len(lines))
	if len(lines) > 1000 {
		skipAlternatedata = true
	}
	for i, line := range lines {
		if skipAlternatedata {
			if i%2 == 0 {
				continue
			}
		}
		words := strings.Fields(line)
		for i, word := range words {
			switch i {
			case 5: //6th column (6th for zero indexed table) is RSS
				switch {
				case strings.Contains(word, "g"):
					word = strings.TrimSuffix(word, "g")
					wordfloat, _ := strconv.ParseFloat(word, 64)
					word = fmt.Sprintf("%0.2f", wordfloat*1024)

				case strings.Contains(word, "m"):
					word = strings.TrimSuffix(word, "m")

				default:
					wordfloat, _ := strconv.ParseFloat(word, 64)
					word = fmt.Sprintf("%0.2f", wordfloat/1024)
				}
				//fmt.Printf("memory = %s\n", word)
				topvalsMem = append(topvalsMem, word)

			case 8: //8th colume (9th for zero indexed table) is CPU
				//fmt.Printf("cpu = %s\n", word)
				topvalCPU = append(topvalCPU, word)
			}
		}
	}

	//fmt.Printf("Memory usage = %s\n", strings.Join(topvalsMem, ","))
	chartFormatStr := fmt.Sprintf(`{"type":"line","series":[{"name":"\"MB\"","values":"{%s}"}],"title":{"name":"Memory Usage"}}`,
		strings.Join(topvalsMem, ","))
	err = xlsx.AddChart(sheet, cellMemory, chartFormatStr)

	chartFormatStr = fmt.Sprintf(`{"type":"line","series":[{"name":"\"Cpu Percent\"","values":"{%s}"}],"title":{"name":"CPU Usage"}}`,
		strings.Join(topvalCPU, ","))

	err = xlsx.AddChart(sheet, cellCPU, chartFormatStr)

	return nil
}

func plotDU(xlsx *excelize.File, topfilename string, sheet string, cell string) error {

	var duvals []string
	var wordfloat float64
	var skipAlternatedata bool

	duData, err := ioutil.ReadFile(topfilename)
	if err != nil {
		return err
	}
	duDataStr := string(duData)

	lines := strings.Split(duDataStr, "\n")
	if len(lines) > 1000 {
		skipAlternatedata = true
	}
	fmt.Printf("Number of du lines processing = %d\n", len(lines))
	for i, line := range lines {
		if skipAlternatedata {
			if i%2 == 0 {
				continue
			}
		}
		words := strings.Fields(line)
		if len(words) > 0 {
			word := words[0]
			switch {
			case strings.Contains(word, "G"):
				word = strings.TrimSuffix(word, "G")
				wordfloat, _ = strconv.ParseFloat(word, 64)
				word = fmt.Sprintf("%d", uint(wordfloat*1024*1024))

			case strings.Contains(word, "M"):
				word = strings.TrimSuffix(word, "M")
				wordfloat, _ = strconv.ParseFloat(word, 64)
				word = fmt.Sprintf("%d", uint(wordfloat*1024))

			case strings.Contains(word, "K"):
				word = strings.TrimSuffix(word, "K")
				wordfloat, _ = strconv.ParseFloat(word, 64)
				word = fmt.Sprintf("%d", uint(wordfloat))

			default:
				wordfloat, _ = strconv.ParseFloat(word, 64)
				word = fmt.Sprintf("%d", uint(wordfloat))
			}
			//fmt.Printf("disk = %s\n", word)
			duvals = append(duvals, word)
		}

	}

	//fmt.Printf("Dis utilization {%s}\n", strings.Join(duvals, ","))

	chartFormatStr := fmt.Sprintf(`{"type":"line","series":[{"name":"\"Disk in MB\"","values":"{%s}"}],"title":{"name":"Disk Usage"}}`,
		strings.Join(duvals, ","))

	fmt.Printf("Disutilization chart formater=%s\n", chartFormatStr)
	err = xlsx.AddChart(sheet, cell, chartFormatStr)
	if err != nil {
		fmt.Printf("Error generating Disk Utilization chart %v\n", err)
	}

	return nil
}

func plotSheet(xlsx *excelize.File, sheet string, sub SubTest) error {

	blackBorder := getBlackBorderCenter(xlsx)
	//Create a new sheet with 'sheet' name
	xlsx.NewSheet(sheet)

	//Tabulate throughtput
	x := 1
	y := 1
	xlsx.SetCellValue(sheet, xAxis[x]+yAxis[y-1], "Throughput")
	var xpos, ypos int
	var value float64
	for j := 0; j < len(sub.memout)+1; j++ {
		for i, v := range Operations {
			xpos = x + j
			ypos = y + i
			if j == 0 {
				xlsx.SetCellValue(sheet, xAxis[xpos]+yAxis[ypos], v)
			} else {
				value = getThroughput(sub.memout[j-1], i)
				xlsx.SetCellValue(sheet, xAxis[xpos]+yAxis[ypos], value)
			}
		}
	}
	xlsx.SetCellStyle(sheet, xAxis[x]+yAxis[y], xAxis[xpos]+yAxis[ypos], blackBorder)

	//Tabulate Latency
	//Latency Starts from B20 (row=y=20, column=x=2)
	x = 1
	y = 19
	xlsx.SetCellValue(sheet, xAxis[x]+yAxis[y-1], "Latency")
	for j := 0; j < len(sub.memout)+1; j++ {
		for i, v := range Operations {
			xpos = x + j
			ypos = y + i
			if j == 0 {
				xlsx.SetCellValue(sheet, xAxis[xpos]+yAxis[ypos], v)
			} else {
				value = getLatency(sub.memout[j-1], i)
				xlsx.SetCellValue(sheet, xAxis[xpos]+yAxis[ypos], value)
			}
		}
	}
	xlsx.SetCellStyle(sheet, xAxis[x]+yAxis[y], xAxis[xpos]+yAxis[ypos], blackBorder)

	//Plot top (row=y=2, Column=x=11 && row=y=20, Column=x=11)
	y = 1
	x = 10
	plotTop(xlsx, sub.topfilename, sheet, xAxis[x]+yAxis[y], xAxis[x]+yAxis[y+18])

	//Plot Latency histogram (row=y=2, column=x=20)
	x = 19
	y = 1
	plotHistogram(xlsx, sub.memout, sheet, xAxis[x]+yAxis[y])

	//Plot disk utilization (row=y=20, Column=x=20)
	x = 19
	y = 19
	plotDU(xlsx, sub.dufilename, sheet, xAxis[x]+yAxis[y])
	return nil
}

func getAverage(sub *SubTest, Op int) float64 {

	var avg float64
	for _, m := range sub.memout {
		avg += getThroughput(m, Op)
	}
	return avg / float64(len(sub.memout))
}

//Prints the summary of the report
func (rs *RedisTest) Summary(xlsx *excelize.File, sheet string) {

	borderStyle := getBlackBorderCenter(xlsx)
	var QOperation []string

	//We start by x,y
	x := 1
	y := 1
	//Throughput Table row names
	for i, o := range Operations {
		xlsx.SetCellValue(sheet, xAxis[x]+yAxis[y+i], o)
		QOperation = append(QOperation, fmt.Sprintf(`\" %s \"`, o))
	}

	//throughput Columen names
	x = 2
	y = 0
	for i, s := range rs.Sub {
		xlsx.SetCellValue(sheet, xAxis[x+i]+yAxis[y], s.Name)
	}

	x = 2
	y = 1
	var series_array []string
	for i, s := range rs.Sub {
		var test_array []string
		for j, _ := range Operations {
			value := getAverage(&s, j)
			xlsx.SetCellValue(sheet, xAxis[x+i]+yAxis[y+j], value)
			test_array = append(test_array, fmt.Sprintf("%f", value))
		}
		//series_array = append(series_array, fmt.Sprintf(`{"name":"\" %s \"", "values":"{%s}", "categories":{%s}}`, s.Name, strings.Join(test_array, ","), strings.Join(Operations, ",")))
		series_array = append(series_array, fmt.Sprintf(`{"name":"\" %s \"", "values":"{%s}", "categories":"Sheet1!$B$2:$B$13"}`, s.Name, strings.Join(test_array, ",")))
	}
	xlsx.SetCellStyle(sheet, xAxis[1]+yAxis[0], xAxis[x+len(rs.Sub)-1]+yAxis[y+len(Operations)], borderStyle)

	x = 2
	y = len(Operations) + 5

	//chartFormatStr := fmt.Sprintf(`{"type":"col","series":[%s],"title":{"name":"Throughput"}}`, strings.Join(series_array, ","))
	chartFormatStr := fmt.Sprintf(`{"type":"col","series":[%s], "title":{"name":"Throughput"}, "dimension":{"width":1280,"height":480}}`, strings.Join(series_array, ","))
	fmt.Printf("format chart=%s\n", chartFormatStr)
	err := xlsx.AddChart(sheet, xAxis[x]+yAxis[y], chartFormatStr)
	if err != nil {
		fmt.Printf("summary throughput chart %v\n", err)
	}

}
