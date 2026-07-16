import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { Package, Truck, Users, AlertTriangle, Plus, ChevronDown, ChevronRight, Search, X, Check, ArrowRight, Warehouse, Loader2, PackageCheck, PackageX, Upload, Clock, UserSearch, Edit3, PackageSearch, Factory, ClipboardList, FileSpreadsheet, TrendingDown, Lock, Trash2, KeyRound, Inbox, LogOut, DownloadCloud, UploadCloud, RotateCcw } from 'lucide-react';
import * as XLSX from 'xlsx';
import { storageGet, storageSet, setAdminToken, onUnauthorizedRequest } from './storage.js';
import { PortalHome } from './Portal.jsx';

const TRANSLATION = {"Defensa frontal completa de frente con malla (Chevrolet 03-07 Diesel)Peso 172 lbs": "FBMC 03-07", "Defensa frontal completa de frente con malla (Chevrolet 07.5-10 Diesel)Peso 172 lbs": "FBMC 7.5-10", "Defensa frontal completa de frente con malla (Chevrolet 11-14 Diesel)Peso 172 lbs": "FBMC 11-14", "Defensa frontal completa de frente con malla (Chevrolet 15-19 Diesel)Peso 172 lbs": "FBMC 15-19", "Defensa frontal completa de frente con malla (Chevrolet 20-23 Diesel) Peso 206.8 Lbs.": "FBMC 20-23", "Defensa frontal completa de frente con malla (Chevrolet 24-25 Diesel) Peso 206.8 Lbs.": "FBMC 24-25", "Defensa frontal completa de frente con malla (Chevrolet 26- Diesel) Peso 206.8 Lbs.": "FBMC 26-", "Defensa frontal completa de frente con malla (Chevrolet 99-02 Diesel)Peso 172 lbs.": "FBMC 99-02", "Defensa frontal completa de frente con malla (Dodge 03-05 Diesel)Peso 172 lbs.": "FBMD 03-05", "Defensa frontal completa de frente con malla (Dodge 06-09 Diesel) peso 172 lbs": "FBMD 06-09", "Defensa frontal completa de frente con malla (Dodge 10-18 Diesel)Peso 172 lbs": "FBMD 10-18", "Defensa frontal completa de frente con malla (Dodge 19-24 Diesel)Peso 238 lbs": "FBMD 19-24", "Defensa frontal completa de frente con malla (Dodge 95-02 Diesel)Peso 172 lbs": "FBMD 95-02", "Defensa frontal completa de frente con malla (Ford 05-07 Diesel)Peso 172 lbs": "FBMF 05-07", "Defensa frontal completa de frente con malla (Ford 08-10 Diesel)Peso 172 lbs": "FBMF 08-10", "Defensa frontal completa de frente con malla (Ford 11-16 F-350 y F-450 Diesel ) Peso 172 lbs.": "FBMF 11-16", "Defensa frontal completa de frente con malla (Ford 17-22 Diesel ) Peso 236 lbs.": "FBMF 17-22", "Defensa frontal completa de frente con malla (Ford 23-25 Diesel ) Peso 236 lbs.": "FBMF 23-26", "Defensa frontal completa de frente con malla (Ford 94-98 Diesel)": "FBMF 94-98", "Defensa frontal completa de frente con malla (Ford 99-04 Diesel)Peso 172 lbs": "FBMF 99-04", "Defensa frontal completa de frente con malla (GMC 03-07 Diesel) Peso 172 Lbs.": "FBMG 03-07", "Defensa frontal completa de frente con malla (GMC 11-14 Diesel) peso 172 lbs": "FBMG 11-14", "Defensa frontal completa de frente con malla (GMC 15-19 Diesel) Peso 172 Lbs.": "FBMG 15-19", "Defensa frontal completa de frente con malla (GMC 20-22 Diesel) Peso 172 Lbs.": "FBMG 20-22", "Defensa frontal completa de frente con malla (GMC 7.5-10 Diesel)Peso 172 lbs": "FBMG 7.5-10", "Defensa frontal completa de frente con malla (GMC 99-02 Diesel)": "FBMG 99-02", "Defensa trasera (Chevy 18-26 Diesel) peso 115 lbs.": "RBCGD 18-26", "Defensa trasera (Chevy y GMC 11-18 Diesel) peso 115 lbs.": "RBCGD 11-18", "Defensa trasera (Chevy y GMC 03-07 Diesel) Peso 115 lbs.": "RBCGD 03-07", "Defensa trasera (Chevy y GMC 07-10 Diesel) peso 115 lbs.": "RBCGD 07-10", "Defensa trasera (Chevy y GMC 20-22 Diesel) peso 115 lbs.": "RBCGD 20-22", "Defensa trasera (Chevy y GMC 99-02 Diesel)": "RBCGD 99-02", "Defensa trasera (Dodge 03-09 Diesel) Peso 115 lbs": "RBDD 03-09", "Defensa trasera (Dodge 24-26 Diesel)": "RBDD 24-26", "Defensa trasera (Dodge 95-02 Diesel) Peso 115 lbs": "RBDD 95-02", "Defensa trasera (Ford 08-16 Diesel)": "RBFD 08-16", "Defensa trasera (Ford 17-22 Diesel)": "RBFD 17-22", "Defensa trasera (Ford 23-26 Diesel)": "RBFD 23-26", "Defensa trasera (Ford 94-98 Diesel)": "RBFD 94-98", "Defensa trasera (Ford 99-07 Diesel) Peso 115 lbs": "RBFD 99-07", "Defensa trasera con luces (Chevy y GMC 99-02 Diesel)": "RBCGD 99-02"}
;
const SEED = {"inventory": [{"sku": "FBMC 03-07", "mx": 15, "us": 33}, {"sku": "FBMC 11-14", "mx": 13, "us": 0}, {"sku": "FBMC 15-19", "mx": 42, "us": 0}, {"sku": "FBMC 20-23", "mx": 26, "us": 35}, {"sku": "FBMC 24-25", "mx": 39, "us": 0}, {"sku": "FBMC 24-26", "mx": 0, "us": 63}, {"sku": "FBMC 26-", "mx": 1, "us": 0}, {"sku": "FBMC 7.5-10", "mx": 20, "us": 8}, {"sku": "FBMC 99-02", "mx": 15, "us": 0}, {"sku": "FBMD 03-05", "mx": 10, "us": 19}, {"sku": "FBMD 06-09", "mx": 25, "us": 0}, {"sku": "FBMD 10-18", "mx": 82, "us": 7}, {"sku": "FBMD 19-24", "mx": 95, "us": 15}, {"sku": "FBMD 95-02", "mx": 24, "us": 6}, {"sku": "FBMF 05-07", "mx": 18, "us": 9}, {"sku": "FBMF 08-10", "mx": 38, "us": 35}, {"sku": "FBMF 11-16", "mx": 63, "us": 4}, {"sku": "FBMF 17-22", "mx": 121, "us": 8}, {"sku": "FBMF 23-25", "mx": 86, "us": 0}, {"sku": "FBMF 94-98", "mx": 36, "us": 0}, {"sku": "FBMF 99-04", "mx": 22, "us": 0}, {"sku": "FBMG 03-07", "mx": 14, "us": 9}, {"sku": "FBMG 11-14", "mx": 18, "us": 0}, {"sku": "FBMG 15-19", "mx": 10, "us": 51}, {"sku": "FBMG 20-22", "mx": 39, "us": 0}, {"sku": "FBMG 20-23", "mx": 0, "us": 24}, {"sku": "FBMG 7.5-10", "mx": 13, "us": 50}, {"sku": "FBMG 99-02", "mx": 13, "us": 7}, {"sku": "RBCGD 03-07", "mx": 5, "us": 0}, {"sku": "RBCGD 07-10", "mx": 18, "us": 0}, {"sku": "RBCGD 11-18", "mx": 13, "us": 0}, {"sku": "RBCGD 18-26", "mx": 1, "us": 0}, {"sku": "RBCGD 20-22", "mx": 9, "us": 0}, {"sku": "RBCGD 20-23", "mx": 0, "us": 16}, {"sku": "RBCGD 99-02", "mx": 23, "us": 6}, {"sku": "RBDD 03-09", "mx": 27, "us": 10}, {"sku": "RBDD 24-26", "mx": 1, "us": 0}, {"sku": "RBDD 95-02", "mx": 24, "us": 21}, {"sku": "RBFD 08-16", "mx": 8, "us": 0}, {"sku": "RBFD 17-22", "mx": 13, "us": 5}, {"sku": "RBFD 23-26", "mx": 2, "us": 0}, {"sku": "RBFD 94-98", "mx": 6, "us": 23}, {"sku": "RBFD 99-07", "mx": 21, "us": 7}], "toolboxItems": [{"item_es": "Caja de Herramientas de 18\"ancho  x 12\" fondo x 18\" alto", "sku": null, "qty": 10}, {"item_es": "Caja de Herramientas de 24\" ancho  x 12\" fondo x 18\"alto", "sku": null, "qty": 22}, {"item_es": "Caja de Herramientas de 30\" largo  x 20\" ancho x 16\"alto", "sku": null, "qty": 10}], "dealers": [{"name": "2-D Trailer Sales LLC", "origin": "US"}, {"name": "4 Mile Trailers", "origin": "US"}, {"name": "4-W Cattle Company, LLC", "origin": "US"}, {"name": "A Day Trailers, LLC", "origin": "MX"}, {"name": "APC Equipment, Inc.", "origin": "MX"}, {"name": "Accel logistics", "origin": "US"}, {"name": "Ag 98 Trailer Sales", "origin": "MX"}, {"name": "Atkins Richard", "origin": "US"}, {"name": "BT Tire & Ag Sales, LLC", "origin": "MX"}, {"name": "Bailey Auto", "origin": "MX"}, {"name": "Bar Circle W. Sales, LLC", "origin": "US"}, {"name": "Bartlett Cooperative Assn.", "origin": "US"}, {"name": "Battleborn Trailer Sales, LLC", "origin": "MX"}, {"name": "Bell Trailerplex", "origin": "MX"}, {"name": "Better Built Trailers", "origin": "MX"}, {"name": "C & J Traders", "origin": "MX"}, {"name": "California Custom Trailers", "origin": "MX"}, {"name": "Christensen Ranches", "origin": "MX"}, {"name": "Clayton Brothers Farm Trust", "origin": "MX"}, {"name": "Colt Bruegman Trailer Sales, LLC", "origin": "MX"}, {"name": "Cross Trail", "origin": "MX"}, {"name": "Crouch Mesa Trailer Sales, LLC", "origin": "MX"}, {"name": "D & E Sales", "origin": "US"}, {"name": "DH Farm Equipment", "origin": "US"}, {"name": "DP Platinum Star Trailers, LLC", "origin": "US"}, {"name": "Diamond K Sales", "origin": "MX"}, {"name": "Diamond T Metals", "origin": "US"}, {"name": "Diamond T Sales", "origin": "US"}, {"name": "Don Evans Legal Window Tint & Trailers", "origin": "US"}, {"name": "Double S Towing Service, LLC", "origin": "US"}, {"name": "Double Z Service & Supply, LLC", "origin": "MX"}, {"name": "Durham Trailer Ranch, Inc.", "origin": "MX"}, {"name": "Early Trailer Sales, LLC", "origin": "MX"}, {"name": "Elite Trailer Sales & Service, LLC", "origin": "US"}, {"name": "F & S Take Off Parts", "origin": "US"}, {"name": "Frye Farms Trailers", "origin": "MX"}, {"name": "GR Trailers LLC. {Customer}", "origin": "MX"}, {"name": "Gengler Auto, LLC", "origin": "US"}, {"name": "Goodman Ag Supply, Inc.", "origin": "MX"}, {"name": "H 5 Ranch, LLC", "origin": "US"}, {"name": "Heacock Trailers & Truck", "origin": "MX"}, {"name": "Hendrys Ron", "origin": "US"}, {"name": "Hisle Brothers, Inc.", "origin": "US"}, {"name": "Hitchin Post Motors", "origin": "US"}, {"name": "Horsch Trailer Sales", "origin": "US"}, {"name": "J&G Trailers", "origin": "MX"}, {"name": "JME Trailers, LLC", "origin": "MX"}, {"name": "Jason Lewis Automotive", "origin": "MX"}, {"name": "Jim's Motors INC", "origin": "MX"}, {"name": "Ken's Trailer Sales & Repair", "origin": "US"}, {"name": "King Gary", "origin": "MX"}, {"name": "Kohler Trailer Sales", "origin": "MX"}, {"name": "Legion Diesel's LLC", "origin": "US"}, {"name": "Lone Star Trailers", "origin": "MX"}, {"name": "MJ Trailers, LLC", "origin": "US"}, {"name": "McFadden Trailer Sales", "origin": "MX"}, {"name": "McMillan AG Repair & Service", "origin": "US"}, {"name": "Meyer Automotive", "origin": "MX"}, {"name": "Mid-Valley Trailer Sales", "origin": "MX"}, {"name": "Monday Trailers & Equipment", "origin": "US"}, {"name": "Northwest Farm Supply", "origin": "MX"}, {"name": "Open Range, LLC", "origin": "US"}, {"name": "Penner Trailer Sales, LLC", "origin": "MX"}, {"name": "Pinnacle Trailer Sales, LLC", "origin": "MX"}, {"name": "Prime Rate Motors, INC.", "origin": "MX"}, {"name": "Producers CO-OP Trailers", "origin": "MX"}, {"name": "Red Barn Trailers, LLC", "origin": "MX"}, {"name": "Redline Trailers and More", "origin": "US"}, {"name": "Riverside Boot & Saddle Blackfoot", "origin": "MX"}, {"name": "Riverside Trailers Caldwell", "origin": "MX"}, {"name": "Riverside Trailers Jerome", "origin": "MX"}, {"name": "Rod's Auto Sales LLC", "origin": "US"}, {"name": "Route 66 Trailer Sales", "origin": "US"}, {"name": "Royer John", "origin": "US"}, {"name": "S & S Motors LLC", "origin": "US"}, {"name": "Scranton Truck & Trailer", "origin": "MX"}, {"name": "Shalom Trailers, INC", "origin": "MX"}, {"name": "Springfield Trailers, INC.", "origin": "US"}, {"name": "Superior Steel Sales, LLC", "origin": "US"}, {"name": "Superior Trailer Sales", "origin": "US"}, {"name": "T & T Quality Building", "origin": "US"}, {"name": "T & T Trailer Sales LLC", "origin": "MX"}, {"name": "TIP TOP TRAILERS", "origin": "MX"}, {"name": "Temco Mfg. Inc.", "origin": "MX"}, {"name": "The Part Xperts", "origin": "MX"}, {"name": "The Truck Shop, LLC", "origin": "US"}, {"name": "Timberline Truck & Trailer", "origin": "US"}, {"name": "Trail Dust Trailers INC", "origin": "MX"}, {"name": "Traileros", "origin": "MX"}, {"name": "Trailquip Plus, LLC", "origin": "US"}, {"name": "Travln Toys", "origin": "MX"}, {"name": "Tri-Star Fleet Sales", "origin": "MX"}, {"name": "Triple M Trailers", "origin": "MX"}, {"name": "Tru-Trailers, Inc.", "origin": "MX"}, {"name": "Truck Country", "origin": "MX"}, {"name": "Unlimited Genetics, Inc.", "origin": "US"}, {"name": "V-Bar Trailer Sales", "origin": "MX"}, {"name": "Walker Trailer And Equipment, LLC", "origin": "US"}, {"name": "West Luke", "origin": "US"}, {"name": "West Texas Trailers & Equipment", "origin": "US"}, {"name": "Wilcoxson Welding", "origin": "US"}, {"name": "Wolfe Auto Service, LLC", "origin": "US"}, {"name": "XB Trailer Sales", "origin": "MX"}, {"name": "YBell Ranch Supply", "origin": "MX"}, {"name": "Zona Trailer Sales, LLC", "origin": "MX"}], "orders": [{"id": "O1", "sku": "BM-FBMC 03-07C", "date": "06/19/2026", "due": "06/19/2026", "num": "39369", "po": "", "dealer": "Cash", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O2", "sku": "FBMC 03-07", "date": "02/27/2026", "due": "03/29/2026", "num": "37987", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O3", "sku": "FBMC 03-07", "date": "07/03/2026", "due": "06/05/2026", "num": "38078", "po": "", "dealer": "Legion Diesel's LLC", "qty": 8, "invoiced": 7, "backordered": 1}, {"id": "O4", "sku": "FBMC 03-07", "date": "04/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 5, "backordered": 5}, {"id": "O5", "sku": "FBMC 03-07", "date": "04/20/2026", "due": "05/20/2026", "num": "38614", "po": "", "dealer": "Christensen Ranches", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O6", "sku": "FBMC 03-07", "date": "12/05/2026", "due": "11/06/2026", "num": "38935", "po": "", "dealer": "Horsch Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O7", "sku": "FBMC 03-07", "date": "5/29/2026", "due": "06/28/2026", "num": "39120", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O8", "sku": "FBMC 03-07", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O9", "sku": "FBMC 11-14", "date": "02/06/2026", "due": "02/07/2026", "num": "39145", "po": "", "dealer": "S & S Motors LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O10", "sku": "FBMC 11-14", "date": "08/06/2026", "due": "08/07/2026", "num": "39209", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O11", "sku": "FBMC 11-14", "date": "09/06/2026", "due": "09/07/2026", "num": "39239", "po": "", "dealer": "Trailquip Plus, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O12", "sku": "FBMC 11-14", "date": "6/17/2026", "due": "7/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O13", "sku": "FBMC 15-19", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O14", "sku": "FBMC 15-19", "date": "06/02/2026", "due": "10/03/2026", "num": "37729", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O15", "sku": "FBMC 15-19", "date": "04/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 9, "backordered": 1}, {"id": "O16", "sku": "FBMC 15-19", "date": "04/14/2026", "due": "05/14/2026", "num": "38529", "po": "", "dealer": "Open Range, LLC", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O17", "sku": "FBMC 15-19", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O18", "sku": "FBMC 15-19", "date": "6/18/2026", "due": "07/18/2026", "num": "39348", "po": "21257", "dealer": "Temco Mfg. Inc.", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O19", "sku": "FBMC 15-19", "date": "6/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O20", "sku": "FBMC 20-23", "date": "12/06/2025", "due": "12/07/2025", "num": "34882", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 9, "backordered": 1}, {"id": "O21", "sku": "FBMC 20-23", "date": "08/15/2025", "due": "09/14/2025", "num": "35640", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O22", "sku": "FBMC 20-23", "date": "06/01/2026", "due": "05/02/2026", "num": "37279", "po": "Dean", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O23", "sku": "FBMC 20-23", "date": "07/03/2026", "due": "06/05/2026", "num": "38078", "po": "", "dealer": "Legion Diesel's LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O24", "sku": "FBMC 20-23", "date": "01/07/2026", "due": "10/08/2026", "num": "39495", "po": "", "dealer": "BT Tire & Ag Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O25", "sku": "FBMC 24-26", "date": "12/06/2025", "due": "12/07/2025", "num": "34882", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 0, "backordered": 10}, {"id": "O26", "sku": "FBMC 24-26", "date": "03/11/2025", "due": "03/12/2025", "num": "36637", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O27", "sku": "FBMC 24-26", "date": "11/03/2026", "due": "10/04/2026", "num": "38112", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O28", "sku": "FBMC 24-26", "date": "11/03/2026", "due": "10/04/2026", "num": "38116", "po": "", "dealer": "JME Trailers, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O29", "sku": "FBMC 24-26", "date": "4/20/2026", "due": "5/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O30", "sku": "FBMC 24-26", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O31", "sku": "FBMC 24-26", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O32", "sku": "FBMC 24-26", "date": "09/06/2026", "due": "09/07/2026", "num": "39244", "po": "", "dealer": "Kohler Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O33", "sku": "FBMC 24-26", "date": "6/24/2026", "due": "07/24/2026", "num": "39419", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O34", "sku": "FBMC 7.5-10", "date": "11/11/2025", "due": "10/12/2025", "num": "36752", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O35", "sku": "FBMC 7.5-10", "date": "07/01/2026", "due": "06/02/2026", "num": "37289", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 7, "backordered": 3}, {"id": "O36", "sku": "FBMC 7.5-10", "date": "3/31/2026", "due": "04/30/2026", "num": "38376", "po": "", "dealer": "Gengler Auto, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O37", "sku": "FBMC 7.5-10", "date": "4/20/2026", "due": "05/20/2026", "num": "38618", "po": "21094", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O38", "sku": "FBMC 7.5-10", "date": "04/22/2026", "due": "5/22/2026", "num": "38690", "po": "21102", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O39", "sku": "FBMC 7.5-10", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O40", "sku": "FBMC 7.5-10", "date": "6/29/2026", "due": "7/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O41", "sku": "FBMC 99-02", "date": "09/06/2026", "due": "09/07/2026", "num": "39234", "po": "", "dealer": "West Luke", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O42", "sku": "FBMC 99-02", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O43", "sku": "FBMD 03-05", "date": "1/26/2026", "due": "2/25/2026", "num": "37560", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 3, "backordered": 2}, {"id": "O44", "sku": "FBMD 03-05", "date": "08/06/2026", "due": "08/07/2026", "num": "39216", "po": "", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O45", "sku": "FBMD 03-05", "date": "12/06/2026", "due": "10/07/2026", "num": "39288", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O46", "sku": "FBMD 03-05", "date": "6/29/2026", "due": "7/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O47", "sku": "FBMD 03-05", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O48", "sku": "FBMD 06-09", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O49", "sku": "FBMD 06-09", "date": "03/16/2026", "due": "10/04/2026", "num": "38194", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O50", "sku": "FBMD 06-09", "date": "03/18/2026", "due": "04/17/2026", "num": "38232", "po": "13025", "dealer": "Diamond K Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O51", "sku": "FBMD 06-09", "date": "03/31/2026", "due": "04/30/2026", "num": "38372", "po": "", "dealer": "S & S Motors LLC", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O52", "sku": "FBMD 06-09", "date": "04/14/2026", "due": "05/14/2026", "num": "38545", "po": "", "dealer": "West Luke", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O53", "sku": "FBMD 06-09", "date": "12/06/2026", "due": "12/07/2026", "num": "39289", "po": "", "dealer": "S & S Motors LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O54", "sku": "FBMD 06-09", "date": "06/23/2026", "due": "07/23/2026", "num": "39406", "po": "", "dealer": "Lone Star Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O55", "sku": "FBMD 06-09", "date": "06/30/2026", "due": "7/30/2026", "num": "39467", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O56", "sku": "FBMD 10-18", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O57", "sku": "FBMD 10-18", "date": "03/25/2026", "due": "4/24/2026", "num": "38329", "po": "21037", "dealer": "Temco Mfg. Inc.", "qty": 4, "invoiced": 3, "backordered": 1}, {"id": "O58", "sku": "FBMD 10-18", "date": "06/04/2026", "due": "06/05/2026", "num": "38432", "po": "", "dealer": "S & S Motors LLC", "qty": 10, "invoiced": 7, "backordered": 3}, {"id": "O59", "sku": "FBMD 10-18", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O60", "sku": "FBMD 10-18", "date": "08/06/2026", "due": "08/07/2026", "num": "39215", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O61", "sku": "FBMD 10-18", "date": "6/17/2026", "due": "7/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O62", "sku": "FBMD 10-18", "date": "6/17/2026", "due": "10/07/2026", "num": "39330", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O63", "sku": "FBMD 10-18", "date": "6/22/2026", "due": "06/22/2026", "num": "39374", "po": "", "dealer": "King Gary", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O64", "sku": "FBMD 10-18", "date": "06/22/2026", "due": "10/07/2026", "num": "39379", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O65", "sku": "FBMD 10-18", "date": "06/29/2026", "due": "07/29/2026", "num": "39459", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O66", "sku": "FBMD 10-18", "date": "6/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O67", "sku": "FBMD 10-18", "date": "01/07/2026", "due": "07/31/2026", "num": "39486", "po": "", "dealer": "Early Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O68", "sku": "FBMD 10-18", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O69", "sku": "FBMD 19-24", "date": "12/30/2025", "due": "01/29/2026", "num": "37207", "po": "", "dealer": "Legion Diesel's LLC", "qty": 15, "invoiced": 10, "backordered": 5}, {"id": "O70", "sku": "FBMD 19-24", "date": "07/01/2026", "due": "06/02/2026", "num": "37289", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 5, "backordered": 5}, {"id": "O71", "sku": "FBMD 19-24", "date": "04/03/2026", "due": "03/04/2026", "num": "38045", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O72", "sku": "FBMD 19-24", "date": "06/03/2026", "due": "05/04/2026", "num": "38073", "po": "3528", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O73", "sku": "FBMD 19-24", "date": "03/17/2026", "due": "4/16/2026", "num": "38209", "po": "3543", "dealer": "Truck Country", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O74", "sku": "FBMD 19-24", "date": "03/31/2026", "due": "4/30/2026", "num": "38377", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O75", "sku": "FBMD 19-24", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O76", "sku": "FBMD 19-24", "date": "10/06/2026", "due": "10/07/2026", "num": "39258", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O77", "sku": "FBMD 19-24", "date": "10/06/2026", "due": "10/07/2026", "num": "39258", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O78", "sku": "FBMD 19-24", "date": "06/17/2026", "due": "07/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 6, "invoiced": 2, "backordered": 4}, {"id": "O79", "sku": "FBMD 19-24", "date": "06/22/2026", "due": "06/22/2026", "num": "39374", "po": "", "dealer": "King Gary", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O80", "sku": "FBMD 19-24", "date": "06/24/2026", "due": "07/24/2026", "num": "39419", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O81", "sku": "FBMD 19-24", "date": "06/25/2026", "due": "10/07/2026", "num": "39428", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O82", "sku": "FBMD 19-24", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O83", "sku": "FBMD 25-26", "date": "02/12/2025", "due": "01/01/2026", "num": "36978", "po": "", "dealer": "S & S Motors LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O84", "sku": "FBMD 25-26", "date": "12/23/2025", "due": "01/22/2026", "num": "37186", "po": "", "dealer": "Diamond T Metals", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O85", "sku": "FBMD 25-26", "date": "02/01/2026", "due": "01/02/2026", "num": "37240", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O86", "sku": "FBMD 25-26", "date": "02/02/2026", "due": "04/03/2026", "num": "37637", "po": "3471", "dealer": "Truck Country", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O87", "sku": "FBMD 25-26", "date": "03/02/2026", "due": "05/03/2026", "num": "37655", "po": "", "dealer": "Kohler Trailer Sales", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O88", "sku": "FBMD 25-26", "date": "05/02/2026", "due": "07/03/2026", "num": "37707", "po": "", "dealer": "Route 66 Trailer Sales", "qty": 8, "invoiced": 0, "backordered": 8}, {"id": "O89", "sku": "FBMD 25-26", "date": "12/02/2026", "due": "03/14/2026", "num": "37802", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O90", "sku": "FBMD 25-26", "date": "06/03/2026", "due": "05/04/2026", "num": "38074", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O91", "sku": "FBMD 25-26", "date": "12/03/2026", "due": "11/04/2026", "num": "38128", "po": "", "dealer": "Texoma Trailers, Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O92", "sku": "FBMD 25-26", "date": "3/23/2026", "due": "07/04/2026", "num": "38296", "po": "", "dealer": "XB Trailer Sales", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O93", "sku": "FBMD 25-26", "date": "4/14/2026", "due": "5/14/2026", "num": "38533", "po": "15-5564", "dealer": "Trailquip Plus, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O94", "sku": "FBMD 25-26", "date": "4/20/2026", "due": "5/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O95", "sku": "FBMD 25-26", "date": "4/30/2026", "due": "05/30/2026", "num": "38817", "po": "", "dealer": "Hitchin Post Motors", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O96", "sku": "FBMD 25-26", "date": "05/05/2026", "due": "04/06/2026", "num": "38859", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O97", "sku": "FBMD 25-26", "date": "11/05/2026", "due": "10/06/2026", "num": "38931", "po": "", "dealer": "D & E Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O98", "sku": "FBMD 25-26", "date": "6/17/2026", "due": "7/17/2026", "num": "39328", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 32, "invoiced": 0, "backordered": 32}, {"id": "O99", "sku": "FBMD 25-26", "date": "6/18/2026", "due": "07/18/2026", "num": "39347", "po": "", "dealer": "Lone Star Trailers", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O100", "sku": "FBMD 25-26", "date": "6/18/2026", "due": "7/18/2026", "num": "39348", "po": "21257", "dealer": "Temco Mfg. Inc.", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O101", "sku": "FBMD 25-26", "date": "6/19/2026", "due": "7/19/2026", "num": "39359", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O102", "sku": "FBMD 25-26", "date": "6/19/2026", "due": "7/19/2026", "num": "39360", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 10, "invoiced": 0, "backordered": 10}, {"id": "O103", "sku": "FBMD 25-26", "date": "6/19/2026", "due": "7/19/2026", "num": "39361", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 0, "backordered": 10}, {"id": "O104", "sku": "FBMD 25-26", "date": "06/19/2026", "due": "07/19/2026", "num": "39362", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O105", "sku": "FBMD 25-26", "date": "06/19/2026", "due": "07/19/2026", "num": "39363", "po": "", "dealer": "D & E Sales", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O106", "sku": "FBMD 25-26", "date": "06/19/2026", "due": "06/19/2026", "num": "39365", "po": "", "dealer": "RZS, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O107", "sku": "FBMD 25-26", "date": "06/22/2026", "due": "07/22/2026", "num": "39371", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O108", "sku": "FBMD 25-26", "date": "06/22/2026", "due": "10/07/2026", "num": "39377", "po": "", "dealer": "JME Trailers, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O109", "sku": "FBMD 25-26", "date": "06/22/2026", "due": "10/07/2026", "num": "39388", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O110", "sku": "FBMD 25-26", "date": "6/23/2026", "due": "6/23/2026", "num": "39398", "po": "", "dealer": "King Gary", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O111", "sku": "FBMD 25-26", "date": "06/25/2026", "due": "7/25/2026", "num": "39433", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O112", "sku": "FBMD 25-26", "date": "06/25/2026", "due": "07/25/2026", "num": "39434", "po": "", "dealer": "Bar Circle W. Sales of Texas, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O113", "sku": "FBMD 95-02", "date": "4/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O114", "sku": "FBMD 95-02", "date": "4/15/2026", "due": "4/15/2026", "num": "38556", "po": "", "dealer": "Route 66 Trailer Sales", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O115", "sku": "FBMD 95-02", "date": "12/05/2026", "due": "11/06/2026", "num": "38934", "po": "21148", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O116", "sku": "FBMD 95-02", "date": "5/27/2026", "due": "06/26/2026", "num": "39087", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O117", "sku": "FBMD 95-02", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O118", "sku": "FBMF 05-07", "date": "11/24/2025", "due": "12/24/2025", "num": "36901", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O119", "sku": "FBMF 05-07", "date": "12/18/2025", "due": "01/17/2026", "num": "37170", "po": "", "dealer": "Gengler Auto, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O120", "sku": "FBMF 05-07", "date": "12/05/2026", "due": "11/06/2026", "num": "38935", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O121", "sku": "FBMF 08-10", "date": "07/03/2026", "due": "06/05/2026", "num": "38078", "po": "", "dealer": "Legion Diesel's LLC", "qty": 5, "invoiced": 4, "backordered": 1}, {"id": "O122", "sku": "FBMF 11-16", "date": "1/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O123", "sku": "FBMF 11-16", "date": "3/14/2026", "due": "04/13/2026", "num": "38179", "po": "13025", "dealer": "Diamond K Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O124", "sku": "FBMF 11-16", "date": "3/20/2026", "due": "10/04/2026", "num": "38272", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O125", "sku": "FBMF 11-16", "date": "06/04/2026", "due": "06/05/2026", "num": "38412", "po": "", "dealer": "JAR Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O126", "sku": "FBMF 11-16", "date": "4/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 2, "backordered": 8}, {"id": "O127", "sku": "FBMF 11-16", "date": "05/22/2026", "due": "06/21/2026", "num": "39054", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O128", "sku": "FBMF 11-16", "date": "02/06/2026", "due": "02/07/2026", "num": "39137", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O129", "sku": "FBMF 11-16", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O130", "sku": "FBMF 11-16", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O131", "sku": "FBMF 11-16", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O132", "sku": "FBMF 11-16", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O133", "sku": "FBMF 17-22", "date": "07/15/2025", "due": "08/14/2025", "num": "35211", "po": "", "dealer": "The Part Xperts", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O134", "sku": "FBMF 17-22", "date": "01/28/2026", "due": "02/27/2026", "num": "37584", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O135", "sku": "FBMF 17-22", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 6, "invoiced": 1, "backordered": 5}, {"id": "O136", "sku": "FBMF 17-22", "date": "02/02/2026", "due": "04/03/2026", "num": "37640", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O137", "sku": "FBMF 17-22", "date": "02/16/2026", "due": "03/18/2026", "num": "37830", "po": "", "dealer": "Christensen Ranches", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O138", "sku": "FBMF 17-22", "date": "03/16/2026", "due": "04/15/2026", "num": "38183", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O139", "sku": "FBMF 17-22", "date": "3/30/2026", "due": "10/04/2026", "num": "38371", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O140", "sku": "FBMF 17-22", "date": "4/20/2026", "due": "05/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O141", "sku": "FBMF 17-22", "date": "02/06/2026", "due": "02/07/2026", "num": "39137", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 8, "invoiced": 0, "backordered": 8}, {"id": "O142", "sku": "FBMF 17-22", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O143", "sku": "FBMF 17-22", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O144", "sku": "FBMF 17-22", "date": "06/17/2026", "due": "07/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 7, "invoiced": 6, "backordered": 1}, {"id": "O145", "sku": "FBMF 17-22", "date": "06/25/2026", "due": "07/25/2026", "num": "39433", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O146", "sku": "FBMF 17-22", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O147", "sku": "FBMF 23-26", "date": "11/25/2025", "due": "12/25/2025", "num": "36937", "po": "", "dealer": "Legion Diesel's LLC", "qty": 10, "invoiced": 5, "backordered": 5}, {"id": "O148", "sku": "FBMF 23-26", "date": "07/01/2026", "due": "06/02/2026", "num": "37289", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 1, "backordered": 9}, {"id": "O149", "sku": "FBMF 23-26", "date": "11/03/2026", "due": "10/04/2026", "num": "38112", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 5, "invoiced": 2, "backordered": 3}, {"id": "O150", "sku": "FBMF 23-26", "date": "03/31/2026", "due": "04/30/2026", "num": "38377", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O151", "sku": "FBMF 23-26", "date": "04/13/2026", "due": "5/13/2026", "num": "38507", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O152", "sku": "FBMF 23-26", "date": "04/15/2026", "due": "5/15/2026", "num": "38555", "po": "", "dealer": "Horsch Trailer Sales", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O153", "sku": "FBMF 23-26", "date": "04/20/2026", "due": "5/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O154", "sku": "FBMF 23-26", "date": "12/05/2026", "due": "11/06/2026", "num": "38943", "po": "21149", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O155", "sku": "FBMF 23-26", "date": "05/18/2026", "due": "06/17/2026", "num": "39001", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 20, "invoiced": 0, "backordered": 20}, {"id": "O156", "sku": "FBMF 23-26", "date": "5/29/2026", "due": "6/28/2026", "num": "39120", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O157", "sku": "FBMF 23-26", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O158", "sku": "FBMF 23-26", "date": "10/06/2026", "due": "10/07/2026", "num": "39255", "po": "", "dealer": "S & S Motors LLC", "qty": 7, "invoiced": 3, "backordered": 4}, {"id": "O159", "sku": "FBMF 23-26", "date": "6/17/2026", "due": "07/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O160", "sku": "FBMF 23-26", "date": "6/25/2026", "due": "07/25/2026", "num": "39433", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O161", "sku": "FBMF 23-26", "date": "6/29/2026", "due": "07/29/2026", "num": "39453", "po": "1630", "dealer": "Diamond T Metals", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O162", "sku": "FBMF 23-26", "date": "06/30/2026", "due": "10/07/2026", "num": "39473", "po": "", "dealer": "Hitchin Post Motors", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O163", "sku": "FBMF 23-26", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O164", "sku": "FBMF 94-98", "date": "03/25/2026", "due": "04/24/2026", "num": "38329", "po": "21037", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O165", "sku": "FBMF 94-98", "date": "04/29/2026", "due": "10/05/2026", "num": "38792", "po": "", "dealer": "DH Farm Equipment", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O166", "sku": "FBMF 94-98", "date": "06/17/2026", "due": "7/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O167", "sku": "FBMF 94-98", "date": "01/07/2026", "due": "7/31/2026", "num": "39487", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O168", "sku": "FBMF 99-04", "date": "11/11/2025", "due": "10/12/2025", "num": "36752", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O169", "sku": "FBMF 99-04", "date": "01/16/2026", "due": "2/15/2026", "num": "37415", "po": "", "dealer": "S & S Motors LLC", "qty": 7, "invoiced": 5, "backordered": 2}, {"id": "O170", "sku": "FBMF 99-04", "date": "02/27/2026", "due": "03/29/2026", "num": "37987", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O171", "sku": "FBMF 99-04", "date": "03/13/2026", "due": "12/04/2026", "num": "38155", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 9, "backordered": 1}, {"id": "O172", "sku": "FBMF 99-04", "date": "12/05/2026", "due": "11/06/2026", "num": "38935", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O173", "sku": "FBMF 99-04", "date": "05/22/2026", "due": "06/21/2026", "num": "39054", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O174", "sku": "FBMF 99-04", "date": "10/06/2026", "due": "10/07/2026", "num": "39250", "po": "", "dealer": "Better Built Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O175", "sku": "FBMF 99-04", "date": "06/15/2026", "due": "07/15/2026", "num": "39304", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O176", "sku": "FBMF 99-04", "date": "06/29/2026", "due": "07/29/2026", "num": "39459", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O177", "sku": "FBMF 99-04", "date": "01/07/2026", "due": "07/31/2026", "num": "39501", "po": "Abegglen", "dealer": "Ken's Trailer Sales & Repair, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O178", "sku": "FBMG 03-07", "date": "12/02/2026", "due": "03/14/2026", "num": "37796", "po": "", "dealer": "XB Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O179", "sku": "FBMG 03-07", "date": "12/02/2026", "due": "03/14/2026", "num": "37802", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O180", "sku": "FBMG 03-07", "date": "02/16/2026", "due": "03/18/2026", "num": "37844", "po": "", "dealer": "XB Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O181", "sku": "FBMG 03-07", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O182", "sku": "FBMG 03-07", "date": "02/07/2026", "due": "10/08/2026", "num": "39513", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O183", "sku": "FBMG 11-14", "date": "06/03/2026", "due": "05/05/2026", "num": "38060", "po": "", "dealer": "Legion Diesel's LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O184", "sku": "FBMG 11-14", "date": "11/04/2026", "due": "11/05/2026", "num": "38495", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O185", "sku": "FBMG 11-14", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O186", "sku": "FBMG 11-14", "date": "06/22/2026", "due": "6/22/2026", "num": "39380", "po": "", "dealer": "Cash", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O187", "sku": "FBMG 15-19", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O188", "sku": "FBMG 15-19", "date": "04/27/2026", "due": "5/27/2026", "num": "38752", "po": "", "dealer": "S & S Motors LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O189", "sku": "FBMG 15-19", "date": "03/07/2026", "due": "02/08/2026", "num": "39524", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O190", "sku": "FBMG 20-23", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O191", "sku": "FBMG 20-23", "date": "03/07/2026", "due": "10/08/2026", "num": "39522", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O192", "sku": "FBMG 24-26", "date": "08/08/2023", "due": "07/09/2023", "num": "27566", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O193", "sku": "FBMG 24-26", "date": "02/02/2024", "due": "03/03/2024", "num": "29440", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O194", "sku": "FBMG 24-26", "date": "02/28/2024", "due": "2/28/2024", "num": "29722", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O195", "sku": "FBMG 24-26", "date": "2/29/2024", "due": "2/29/2024", "num": "29745", "po": "", "dealer": "Route 66 Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O196", "sku": "FBMG 24-26", "date": "3/13/2024", "due": "12/04/2024", "num": "29909", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O197", "sku": "FBMG 24-26", "date": "02/04/2024", "due": "02/04/2024", "num": "30130", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O198", "sku": "FBMG 24-26", "date": "05/06/2024", "due": "05/06/2024", "num": "30898", "po": "", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O199", "sku": "FBMG 24-26", "date": "11/07/2024", "due": "11/07/2024", "num": "31303", "po": "", "dealer": "West Luke", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O200", "sku": "FBMG 24-26", "date": "7/25/2024", "due": "8/24/2024", "num": "31466", "po": "", "dealer": "Superior Steel Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O201", "sku": "FBMG 24-26", "date": "8/22/2024", "due": "9/21/2024", "num": "32635", "po": "", "dealer": "Legion Diesel's LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O202", "sku": "FBMG 24-26", "date": "12/11/2024", "due": "12/12/2024", "num": "32586", "po": "", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O203", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32662", "po": "Shawn", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O204", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32663", "po": "Donald", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O205", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32664", "po": "Amanda", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O206", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32665", "po": "Was", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O207", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32666", "po": "Kyle", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O208", "sku": "FBMG 24-26", "date": "12/13/2024", "due": "12/01/2025", "num": "32915", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O209", "sku": "FBMG 24-26", "date": "12/16/2024", "due": "01/15/2025", "num": "32924", "po": "", "dealer": "Diamond K Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O210", "sku": "FBMG 24-26", "date": "1/27/2025", "due": "2/26/2025", "num": "33288", "po": "", "dealer": "Diamond K Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O211", "sku": "FBMG 24-26", "date": "10/02/2025", "due": "12/03/2025", "num": "33446", "po": "", "dealer": "Kohler Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O212", "sku": "FBMG 24-26", "date": "03/13/2025", "due": "12/04/2025", "num": "33815", "po": "", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O213", "sku": "FBMG 24-26", "date": "04/21/2025", "due": "5/21/2025", "num": "34285", "po": "", "dealer": "S & S Motors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O214", "sku": "FBMG 24-26", "date": "4/28/2025", "due": "10/05/2025", "num": "34392", "po": "", "dealer": "JME Trailers, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O215", "sku": "FBMG 24-26", "date": "07/07/2025", "due": "06/08/2025", "num": "35130", "po": "", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O216", "sku": "FBMG 24-26", "date": "07/15/2025", "due": "8/14/2025", "num": "35211", "po": "", "dealer": "The Part Xperts", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O217", "sku": "FBMG 24-26", "date": "09/16/2025", "due": "10/31/2025", "num": "36022", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O218", "sku": "FBMG 24-26", "date": "09/16/2025", "due": "10/31/2025", "num": "36022", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O219", "sku": "FBMG 24-26", "date": "12/11/2025", "due": "12/12/2025", "num": "36767", "po": "3363", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O220", "sku": "FBMG 24-26", "date": "12/22/2025", "due": "01/21/2026", "num": "37184", "po": "", "dealer": "H5 Feed & Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O221", "sku": "FBMG 24-26", "date": "06/01/2026", "due": "05/02/2026", "num": "37279", "po": "Dean", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O222", "sku": "FBMG 24-26", "date": "12/01/2026", "due": "11/02/2026", "num": "37361", "po": "3413", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O223", "sku": "FBMG 24-26", "date": "03/02/2026", "due": "05/03/2026", "num": "37655", "po": "", "dealer": "Kohler Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O224", "sku": "FBMG 24-26", "date": "09/02/2026", "due": "11/03/2026", "num": "37761", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O225", "sku": "FBMG 24-26", "date": "02/19/2026", "due": "3/21/2026", "num": "37877", "po": "", "dealer": "H5 Feed & Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O226", "sku": "FBMG 24-26", "date": "04/23/2026", "due": "05/23/2026", "num": "38719", "po": "3616", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O227", "sku": "FBMG 7.5-10", "date": "12/02/2026", "due": "03/14/2026", "num": "37802", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O228", "sku": "FBMG 7.5-10", "date": "2/27/2026", "due": "3/29/2026", "num": "37987", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O229", "sku": "FBMG 7.5-10", "date": "4/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O230", "sku": "FBMG 99-02", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O231", "sku": "RBCGD 11-19", "date": "12/05/2026", "due": "11/06/2026", "num": "38946", "po": "", "dealer": "Open Range, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O232", "sku": "RBCGD 11-19", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O233", "sku": "RBCGD 20-23", "date": "6/22/2026", "due": "06/22/2026", "num": "39374", "po": "", "dealer": "King Gary", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O234", "sku": "RBCGD 20-23", "date": "06/23/2026", "due": "07/23/2026", "num": "39406", "po": "", "dealer": "Lone Star Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O235", "sku": "RBDD 10-23", "date": "02/25/2026", "due": "03/27/2026", "num": "37951", "po": "", "dealer": "Open Range, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O236", "sku": "RBDD 10-23", "date": "02/03/2026", "due": "01/04/2026", "num": "38006", "po": "", "dealer": "West Luke", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O237", "sku": "RBDD 10-23", "date": "03/30/2026", "due": "10/04/2026", "num": "38371", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O238", "sku": "RBDD 10-23", "date": "04/27/2026", "due": "05/27/2026", "num": "38766", "po": "Misenheimer", "dealer": "H5 Feed & Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O239", "sku": "RBDD 10-23", "date": "06/18/2026", "due": "07/18/2026", "num": "39346", "po": "", "dealer": "Scranton Truck & Trailer", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O240", "sku": "RBDD 10-23", "date": "06/23/2026", "due": "07/23/2026", "num": "39406", "po": "", "dealer": "Lone Star Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O241", "sku": "RBDD 10-23", "date": "01/07/2026", "due": "07/31/2026", "num": "39486", "po": "", "dealer": "Early Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O242", "sku": "RBDD 25-26", "date": "09/18/2024", "due": "10/18/2024", "num": "31985", "po": "", "dealer": "D & E Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O243", "sku": "RBDD 25-26", "date": "12/13/2024", "due": "12/01/2025", "num": "32923", "po": "", "dealer": "D & E Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O244", "sku": "RBDD 25-26", "date": "09/16/2025", "due": "10/31/2025", "num": "36022", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O245", "sku": "RBDD 25-26", "date": "06/25/2026", "due": "10/07/2026", "num": "39428", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O246", "sku": "RBDD 95-02", "date": "12/11/2025", "due": "12/12/2025", "num": "36767", "po": "3363", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O247", "sku": "RBFD 08-16", "date": "02/25/2026", "due": "03/27/2026", "num": "37961", "po": "3512", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O248", "sku": "RBFD 08-16", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O249", "sku": "RBFD 17-22", "date": "12/01/2026", "due": "11/02/2026", "num": "37358", "po": "3412", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O250", "sku": "RBFD 17-22", "date": "01/28/2026", "due": "02/27/2026", "num": "37592", "po": "3466", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O251", "sku": "RBFD 23-26", "date": "08/15/2024", "due": "09/14/2024", "num": "31694", "po": "Colby Eli/Cody/Tyler/Dust", "dealer": "Route 66 Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O252", "sku": "RBFD 23-26", "date": "05/02/2025", "due": "07/03/2025", "num": "33386", "po": "", "dealer": "D & E Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O253", "sku": "RBFD 23-26", "date": "05/22/2025", "due": "10/06/2025", "num": "34692", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O254", "sku": "RBFD 23-26", "date": "02/13/2026", "due": "03/15/2026", "num": "37818", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}]};
const LOGO_DATA_URL = 'data:image/jpeg;base64,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';


const LOCATIONS = {
  MX: { key: 'mx', label: 'Mexico Warehouse', short: 'MX', color: '#B23A2E' },
  US: { key: 'us', label: 'GR / US Warehouse', short: 'US', color: '#33546E' },
};

function normName(s) {
  return (s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function formatDateCell(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) {
    const mm = v.getMonth() + 1, dd = v.getDate(), yyyy = v.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
  }
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${parseInt(m[1], 10)}/${parseInt(m[2], 10)}/${m[3]}`;
  return s;
}
function toNum(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string' && v.trim() !== '' && !isNaN(v.trim())) return parseFloat(v.trim());
  return 0;
}

// Parses the QuickBooks-style "open sales orders by item" export: rows with a SKU
// header (col 4) followed by data rows (Type/Date/Due/Num/PO/Name/Qty/Invoiced/Backordered
// in cols 6-14). Mirrors the structure of the original Orders.xlsx export.
function parseOrdersRows(rows) {
  const out = [];

  // Find the header row (contains "Num", "P.O.#", "Type", etc.) and map each
  // column by its label, instead of assuming a fixed position. This way the
  // parser keeps working correctly even if a future export shifts columns
  // around — it always finds "Num" and "P.O.#" wherever they actually are.
  let headerRowIndex = -1;
  let col = { type: 6, date: 7, due: 8, num: 9, po: 10, name: 11, qty: 12, invoiced: 13, backordered: 14 };
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i];
    if (!Array.isArray(row)) continue;
    const numIdx = row.findIndex(c => typeof c === 'string' && c.trim() === 'Num');
    if (numIdx !== -1) {
      headerRowIndex = i;
      const map = {};
      row.forEach((cell, ci) => { if (typeof cell === 'string' && cell.trim()) map[cell.trim()] = ci; });
      col = {
        type: map['Type'] ?? col.type,
        date: map['Date'] ?? col.date,
        due: map['Due Date'] ?? col.due,
        num: map['Num'] ?? col.num,
        po: map['P.O.#'] ?? col.po,
        name: map['Name'] ?? col.name,
        qty: map['Qty'] ?? col.qty,
        invoiced: map['Invoiced'] ?? col.invoiced,
        backordered: map['Backordered'] ?? col.backordered,
      };
      break;
    }
  }
  const startRow = headerRowIndex >= 0 ? headerRowIndex + 1 : 3;

  let currentSku = null;
  rows.forEach((row, i) => {
    if (i < startRow || !Array.isArray(row)) return;
    const skuCell = row[4];
    const typeCell = row[col.type];
    const hasType = typeCell != null && String(typeCell).trim() !== '';
    if (skuCell != null && String(skuCell).trim() !== '' && !hasType) {
      const s = String(skuCell).trim();
      if (/^total/i.test(s)) return;
      currentSku = s;
    } else if (hasType && currentSku) {
      out.push({
        sku: currentSku,
        date: formatDateCell(row[col.date]),
        due: formatDateCell(row[col.due]),
        // "Num" is the real reference number dealers use — that's what we show as PO.
        // "P.O.#" is kept in the num field only for reference; it's almost always blank.
        num: row[col.po] != null ? String(row[col.po]) : '',
        po: row[col.num] != null ? String(row[col.num]) : '',
        dealer: row[col.name] != null ? String(row[col.name]).trim() : '',
        qty: toNum(row[col.qty]),
        invoiced: toNum(row[col.invoiced]),
        backordered: toNum(row[col.backordered]),
      });
    }
  });
  return out;
}

function formatToday() {
  const d = new Date();
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function isoToUS(iso) {
  if (!iso) return formatToday();
  const [y, m, d] = iso.split('-');
  return `${parseInt(m, 10)}/${parseInt(d, 10)}/${y}`;
}
function usToISO(us) {
  const d = parseDate(us);
  if (!d) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseDate(str) {
  if (!str) return null;
  const m = String(str).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3], 10), parseInt(m[1], 10) - 1, parseInt(m[2], 10));
}
function daysSince(str) {
  const d = parseDate(str);
  if (!d) return null;
  return Math.max(0, Math.floor((Date.now() - d.getTime()) / 86400000));
}

// Generic parser for both the Mexico ("Existencia y valor del inventario") export
// and the US/GR stock export. Works on raw rows from XLSX.utils.sheet_to_json(sheet,{header:1}).
function parseInventoryRows(rows, isUS) {
  const matched = [];
  const unmatched = [];
  const skipLabel = /^u\.?\s*med|^art[ií]culo|^solo exist|^l[ií]nea|^al\s+\d|almac[eé]n|grupo remolques|existencia y valor|^total\s+\d/i;
  rows.forEach(row => {
    if (!Array.isArray(row)) return;
    let label = null;
    for (const cell of row) {
      if (typeof cell === 'string') {
        const t = cell.trim();
        if (t.length > 2 && !skipLabel.test(t)) { label = t; break; }
      }
    }
    if (!label) return;
    let qty = null;
    for (let i = row.length - 1; i >= 0; i--) {
      const v = row[i];
      if (typeof v === 'number') { qty = v; break; }
      if (typeof v === 'string' && /^\d+(\.\d+)?$/.test(v.trim())) { qty = parseFloat(v.trim()); break; }
    }
    if (qty === null) return;
    if (isUS) {
      const sku = label.replace(/^total\s+/i, '').trim();
      matched.push({ sku, qty: Math.round(qty), label });
    } else {
      const clean = label.replace(/\s+/g, ' ').trim();
      const sku = TRANSLATION[clean];
      if (sku) matched.push({ sku, qty: Math.round(qty), label: clean });
      else unmatched.push({ label: clean, qty: Math.round(qty) });
    }
  });
  return { matched, unmatched };
}

function ClearButton({ onClick }) {
  return (
    <button onClick={onClick} style={{
      position: 'absolute', right: 6, top: '50%', transform: 'translateY(-50%)',
      border: 'none', background: '#E9E7DE', color: '#5B6470', display: 'flex', alignItems: 'center', justifyContent: 'center',
      width: 17, height: 17, borderRadius: '50%', padding: 0, cursor: 'pointer'
    }}><X size={11} /></button>
  );
}

function Toast({ message, onDone }) {
  useEffect(() => {
    const t = setTimeout(onDone, 2400);
    return () => clearTimeout(t);
  }, [onDone]);
  return (
    <div style={{
      position: 'fixed', bottom: 24, left: '50%', transform: 'translateX(-50%)',
      background: '#1C2126', color: '#F5F3EE', padding: '10px 18px', borderRadius: 8,
      fontSize: 13, fontFamily: "'Inter', sans-serif", boxShadow: '0 8px 24px rgba(0,0,0,0.25)',
      zIndex: 1000, display: 'flex', alignItems: 'center', gap: 8, border: '1px solid #33546E'
    }}>
      <Check size={14} color="#7FC79A" />
      {message}
    </div>
  );
}

function Badge({ children, color, bg, border }) {
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', gap: 3, padding: '1px 6px',
      borderRadius: 4, fontSize: 10, fontWeight: 700, letterSpacing: '0.02em',
      color: color, background: bg, border: `1px solid ${border || bg}`,
      fontFamily: "'IBM Plex Mono', monospace", textTransform: 'uppercase'
    }}>{children}</span>
  );
}

const DATA_VERSION = 5;

function migrateData({ inventory, toolboxItems, dealers, orders }) {
  let inv = inventory.map(r => ({ ...r }));
  let dl = dealers.map(d => ({ ...d }));
  let ord = orders.map(o => ({ ...o }));

  function renameSku(from, to) {
    const idxFrom = inv.findIndex(r => r.sku === from);
    if (idxFrom === -1) return;
    const idxTo = inv.findIndex(r => r.sku === to);
    if (idxTo === -1) {
      inv[idxFrom] = { ...inv[idxFrom], sku: to };
    } else {
      inv[idxTo] = { sku: to, mx: (inv[idxTo].mx || 0) + (inv[idxFrom].mx || 0), us: (inv[idxTo].us || 0) + (inv[idxFrom].us || 0) };
      inv.splice(idxFrom, 1);
    }
  }
  // Correct SKU codes to match the codes actually used on orders
  renameSku('FBMF 23-25', 'FBMF 23-26');
  renameSku('RBCGD 11-18', 'RBCGD 11-19');

  function ensureSku(sku) {
    if (!inv.some(r => r.sku === sku)) inv.push({ sku, mx: 0, us: 0 });
  }
  // These are real SKUs that show up on orders but currently carry no stock
  ['FBMD 25-26', 'FBMG 24-26', 'RBDD 10-23', 'RBDD 25-26'].forEach(ensureSku);

  // BM-FBMC 03-07C isn't a real SKU code - drop the order line entirely
  ord = ord.filter(o => o.sku !== 'BM-FBMC 03-07C');

  // Any dealer origin stored as the old 'GR' code should be 'US' (matches LOCATIONS keys)
  dl = dl.map(d => d.origin === 'GR' ? { ...d, origin: 'US' } : d);

  // Any dealer on an order that isn't in the dealer list yet: add as a US (GR-ships) dealer
  const knownNames = new Set(dl.map(d => normName(d.name)));
  const orderDealerNames = [...new Set(ord.map(o => o.dealer))];
  orderDealerNames.forEach(name => {
    if (!knownNames.has(normName(name))) {
      dl.push({ name, origin: 'US' });
      knownNames.add(normName(name));
    }
  });
  dl.sort((a, b) => a.name.localeCompare(b.name));

  // The QuickBooks import previously read PO # from the "P.O.#" column, but that column
  // is almost always blank — the real reference number lives in "Num". Fix already-imported
  // orders by swapping the two (only orders from the bulk import have a Num value; manually
  // added orders leave Num blank, so they're untouched).
  ord = ord.map(o => (o.num ? { ...o, po: o.num, num: o.po } : o));

  return { inventory: inv, toolboxItems, dealers: dl, orders: ord };
}

function SkuTag({ sku }) {
  return (
    <span style={{
      display: 'inline-block', padding: '2px 7px', borderRadius: 4,
      background: 'linear-gradient(180deg, #2A3038 0%, #1C2126 100%)',
      color: '#F5F3EE', fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5,
      fontWeight: 600, letterSpacing: '0.02em',
      boxShadow: 'inset 0 1px 0 rgba(255,255,255,0.08), 0 1px 2px rgba(0,0,0,0.3)',
      border: '1px solid #0F1215'
    }}>{sku}</span>
  );
}

function Dashboard({ onAdminLogout, adminSession } = {}) {
  const [loading, setLoading] = useState(true);
  const [inventory, setInventory] = useState([]);
  const [toolboxItems, setToolboxItems] = useState([]);
  const [dealers, setDealers] = useState([]);
  const [orders, setOrders] = useState([]);
  const [salesLog, setSalesLog] = useState([]);
  const [productionBatches, setProductionBatches] = useState([]);
  const [productionLog, setProductionLog] = useState([]);
  const [newOrders, setNewOrders] = useState([]);
  const [shippedLog, setShippedLog] = useState([]);
  const [tab, setTab] = useState('production');
  const [toast, setToast] = useState(null);
  const [shipModal, setShipModal] = useState(null);
  const [orderModal, setOrderModal] = useState(false);
  const [dealerModal, setDealerModal] = useState(false);
  const [importModal, setImportModal] = useState(null);
  const [dealerLoginModal, setDealerLoginModal] = useState(null);
  const [editOrderModal, setEditOrderModal] = useState(null);

  useEffect(() => {
    (async () => {
      const [inv, tb, dl, ord, meta, sales, legacyPending, prodLog, batches, incoming, shipped] = await Promise.all([
        storageGet('bumper-inventory', true),
        storageGet('bumper-toolbox', true),
        storageGet('bumper-dealers', true),
        storageGet('bumper-orders', true),
        storageGet('bumper-meta', true),
        storageGet('bumper-sales', true),
        storageGet('bumper-production-pending', true),
        storageGet('bumper-production-log', true),
        storageGet('bumper-production-batches', true),
        storageGet('bumper-new-orders', true),
        storageGet('bumper-shipped-log', true),
      ]);
      let state = (inv && dl && ord)
        ? { inventory: inv, toolboxItems: tb || [], dealers: dl, orders: ord }
        : { inventory: SEED.inventory, toolboxItems: SEED.toolboxItems, dealers: SEED.dealers, orders: SEED.orders };

      const currentVersion = meta?.version || 0;
      if (currentVersion < DATA_VERSION) {
        state = migrateData(state);
        await Promise.all([
          storageSet('bumper-inventory', state.inventory, true),
          storageSet('bumper-toolbox', state.toolboxItems, true),
          storageSet('bumper-dealers', state.dealers, true),
          storageSet('bumper-orders', state.orders, true),
          storageSet('bumper-meta', { version: DATA_VERSION }, true),
        ]);
      } else if (!inv || !dl || !ord) {
        await Promise.all([
          storageSet('bumper-inventory', state.inventory, true),
          storageSet('bumper-toolbox', state.toolboxItems, true),
          storageSet('bumper-dealers', state.dealers, true),
          storageSet('bumper-orders', state.orders, true),
        ]);
      }

      setInventory(state.inventory);
      setToolboxItems(state.toolboxItems);
      setDealers(state.dealers);
      setOrders(state.orders);
      setSalesLog(sales || []);
      setProductionLog(prodLog || []);

      let finalBatches = batches;
      if (!finalBatches) {
        // One-time migration from the old {sku: qty} pending-production map, which had
        // no date info, into dated batches (date left blank for these legacy entries).
        finalBatches = (legacyPending && typeof legacyPending === 'object' && !Array.isArray(legacyPending))
          ? Object.entries(legacyPending).filter(([, q]) => q > 0).map(([sku, qty]) => ({ id: 'PB-legacy-' + sku, sku, qty, date: '' }))
          : [];
        await storageSet('bumper-production-batches', finalBatches, true);
      }
      setProductionBatches(finalBatches);
      setNewOrders(incoming || []);
      setShippedLog(shipped || []);

      setLoading(false);
    })();
  }, []);

  // Keep the app in sync with what everyone else is doing, without needing a manual
  // refresh — quietly re-check the shared data every few seconds and update state if
  // it changed. This never touches the one-time seed/migration logic above.
  useEffect(() => {
    if (loading) return;
    const interval = setInterval(async () => {
      const [inv, tb, dl, ord, sales, batches, prodLog, incoming, shipped] = await Promise.all([
        storageGet('bumper-inventory', true),
        storageGet('bumper-toolbox', true),
        storageGet('bumper-dealers', true),
        storageGet('bumper-orders', true),
        storageGet('bumper-sales', true),
        storageGet('bumper-production-batches', true),
        storageGet('bumper-production-log', true),
        storageGet('bumper-new-orders', true),
        storageGet('bumper-shipped-log', true),
      ]);
      if (inv) setInventory(inv);
      if (tb) setToolboxItems(tb);
      if (dl) setDealers(dl);
      if (ord) setOrders(ord);
      if (sales) setSalesLog(sales);
      if (batches) setProductionBatches(batches);
      if (prodLog) setProductionLog(prodLog);
      if (incoming) setNewOrders(incoming);
      if (shipped) setShippedLog(shipped);
    }, 10000);
    return () => clearInterval(interval);
  }, [loading]);

  const showToast = useCallback((msg) => setToast(msg), []);

  const dealerOriginMap = useMemo(() => {
    const m = {};
    dealers.forEach(d => { m[normName(d.name)] = d.origin; });
    return m;
  }, [dealers]);

  const inventorySkuSet = useMemo(() => new Set(inventory.map(i => i.sku)), [inventory]);

  const ordersEnriched = useMemo(() => {
    return orders.map(o => {
      const origin = dealerOriginMap[normName(o.dealer)];
      const shipFrom = origin || 'MX';
      const dealerKnown = !!origin;
      const skuKnown = inventorySkuSet.has(o.sku);
      return { ...o, shipFrom, dealerKnown, skuKnown };
    });
  }, [orders, dealerOriginMap, inventorySkuSet]);

  const openOrders = useMemo(() => ordersEnriched.filter(o => o.backordered > 0), [ordersEnriched]);

  const dataIssues = useMemo(() => {
    const unknownDealers = [...new Set(openOrders.filter(o => !o.dealerKnown).map(o => o.dealer))];
    const unknownSkus = [...new Set(openOrders.filter(o => !o.skuKnown).map(o => o.sku))];
    return { unknownDealers, unknownSkus };
  }, [openOrders]);

  const demandByLocSku = useMemo(() => {
    const m = {};
    openOrders.forEach(o => {
      const key = o.shipFrom + '|' + o.sku;
      m[key] = (m[key] || 0) + o.backordered;
    });
    return m;
  }, [openOrders]);

  const pendingBySku = useMemo(() => {
    const m = {};
    productionBatches.forEach(b => { m[b.sku] = (m[b.sku] || 0) + b.qty; });
    return m;
  }, [productionBatches]);

  const ordersByLocSku = useMemo(() => {
    const m = {};
    openOrders.forEach(o => {
      const key = o.shipFrom + '|' + o.sku;
      if (!m[key]) m[key] = [];
      m[key].push(o);
    });
    return m;
  }, [openOrders]);

  async function persistInventory(next) {
    setInventory(next);
    await storageSet('bumper-inventory', next, true);
  }
  async function persistOrders(next) {
    setOrders(next);
    await storageSet('bumper-orders', next, true);
  }
  async function persistDealers(next) {
    setDealers(next);
    await storageSet('bumper-dealers', next, true);
  }
  async function persistToolboxItems(next) {
    setToolboxItems(next);
    await storageSet('bumper-toolbox', next, true);
  }
  async function persistSales(next) {
    setSalesLog(next);
    await storageSet('bumper-sales', next, true);
  }
  async function persistProductionBatches(next) {
    setProductionBatches(next);
    await storageSet('bumper-production-batches', next, true);
  }
  async function persistProductionLog(next) {
    setProductionLog(next);
    await storageSet('bumper-production-log', next, true);
  }

  async function persistShippedLog(next) {
    setShippedLog(next);
    await storageSet('bumper-shipped-log', next, true);
  }

  function applyImport(nextInventory, sales, productionReceipts) {
    persistInventory(nextInventory);
    if (sales && sales.length) {
      const withIds = sales.map((s, i) => ({ id: 'S' + Date.now() + '_' + i, ...s }));
      persistSales([...salesLog, ...withIds]);
    }
    let receivedNote = '';
    if (productionReceipts && productionReceipts.length) {
      // Apply each SKU's received units against its own oldest-first batches — same FIFO
      // idea used everywhere else in the app, just applied to production instead of orders.
      let nextBatches = productionBatches.map(b => ({ ...b }));
      const logEntries = [];
      productionReceipts.forEach(r => {
        let remaining = r.qty;
        let applied = 0;
        const skuBatches = nextBatches
          .filter(b => b.sku === r.sku)
          .sort((a, b) => (parseDate(a.date)?.getTime() ?? -Infinity) - (parseDate(b.date)?.getTime() ?? -Infinity));
        for (const b of skuBatches) {
          if (remaining <= 0) break;
          const take = Math.min(b.qty, remaining);
          b.qty -= take;
          remaining -= take;
          applied += take;
        }
        nextBatches = nextBatches.filter(b => b.qty > 0);
        logEntries.push({
          id: 'PR' + Date.now() + '_' + r.sku, type: 'received', sku: r.sku, qty: r.qty,
          appliedToPending: applied, date: r.date, previousQty: r.previousQty, newQty: r.newQty,
        });
      });
      persistProductionBatches(nextBatches);
      persistProductionLog([...productionLog, ...logEntries]);
      const totalApplied = logEntries.reduce((a, b) => a + b.appliedToPending, 0);
      if (totalApplied > 0) receivedNote = ` — ${totalApplied} units credited against production`;
    }
    const soldNote = sales && sales.length ? ` — ${sales.reduce((a, b) => a + b.qty, 0)} units logged as sold` : '';
    showToast(`${LOCATIONS[importModal].label} stock updated from import${soldNote}${receivedNote}`);
    setImportModal(null);
  }

  function saveProductionOrder(items, date) {
    if (!items || items.length === 0) return;
    const orderDate = date || formatToday();
    const newBatches = items.map(({ sku, qty }) => ({ id: 'PB' + Date.now() + '_' + sku, sku, qty, date: orderDate }));
    persistProductionBatches([...productionBatches, ...newBatches]);
    const logEntries = items.map(({ sku, qty }) => ({ id: 'PO' + Date.now() + '_' + sku, type: 'ordered', sku, qty, date: orderDate }));
    persistProductionLog([...productionLog, ...logEntries]);
    showToast(`Pedido de defensas saved for ${orderDate} — ${items.length} model${items.length === 1 ? '' : 's'} sent to production`);
  }

  function updateProductionBatch(id, updates) {
    const next = productionBatches.map(b => b.id === id ? { ...b, ...updates } : b);
    persistProductionBatches(next);
  }

  function deleteProductionBatch(id) {
    persistProductionBatches(productionBatches.filter(b => b.id !== id));
    showToast('Removed that production entry');
  }

  function applyOrdersImport(nextOrders) {
    persistOrders(nextOrders);
    showToast(`Orders replaced from import — ${nextOrders.length} order lines`);
    setImportModal(null);
  }

  function shipOrderQty(order, shipQty) {
    const qty = Math.min(order.backordered, Math.max(0, shipQty));
    if (qty <= 0) return;
    const nextOrders = orders.map(o => o.id === order.id
      ? { ...o, backordered: o.backordered - qty, invoiced: o.invoiced + qty }
      : o);
    persistOrders(nextOrders);
    const entry = {
      id: 'SH' + Date.now(), orderId: order.id, sku: order.sku, dealer: order.dealer,
      po: order.po || '', shipFrom: order.shipFrom, qty, date: formatToday(),
    };
    persistShippedLog([entry, ...shippedLog]);
    showToast(`Shipped ${qty} × ${order.sku} to ${order.dealer}`);
    setShipModal(null);
  }

  function addOrder(items) {
    const list = Array.isArray(items) ? items : [items];
    if (list.length === 0) return;
    let nextId = Math.max(0, ...orders.map(o => parseInt(o.id.slice(1), 10) || 0));
    const created = list.map(item => {
      nextId += 1;
      return { id: 'O' + nextId, invoiced: 0, backordered: item.qty, ...item };
    });
    persistOrders([...created, ...orders]);
    showToast(`Added ${created.length} order line${created.length === 1 ? '' : 's'} for ${list[0].dealer}`);
    setOrderModal(false);
  }

  function updateOrder(id, updates) {
    const next = orders.map(o => o.id === id ? { ...o, ...updates } : o);
    persistOrders(next);
    showToast('Order updated');
  }

  function deleteOrder(id) {
    persistOrders(orders.filter(o => o.id !== id));
    showToast('Order deleted');
  }

  async function persistNewOrders(next) {
    setNewOrders(next);
    await storageSet('bumper-new-orders', next, true);
  }

  function processNewOrders(ids, po) {
    const idSet = new Set(ids);
    const toProcess = newOrders.filter(o => idSet.has(o.id)).map(o => ({ ...o, po: po || o.po }));
    if (toProcess.length === 0) return;
    persistNewOrders(newOrders.filter(o => !idSet.has(o.id)));
    persistOrders([...toProcess, ...orders]);
    showToast(`Processed ${toProcess.length} order line${toProcess.length === 1 ? '' : 's'} for ${toProcess[0].dealer} — PO ${po}`);
  }

  function rejectNewOrders(ids) {
    const idSet = new Set(ids);
    persistNewOrders(newOrders.filter(o => !idSet.has(o.id)));
    showToast('Order request rejected');
  }

  function restoreFromFile(parsed) {
    const d = parsed?.data;
    if (!d) { showToast('That file doesn\'t look like a valid backup'); return; }
    if (d.inventory) persistInventory(d.inventory);
    if (d.toolboxItems) persistToolboxItems(d.toolboxItems);
    if (d.dealers) persistDealers(d.dealers);
    if (d.orders) persistOrders(d.orders);
    if (d.salesLog) persistSales(d.salesLog);
    if (d.productionBatches) persistProductionBatches(d.productionBatches);
    if (d.productionLog) persistProductionLog(d.productionLog);
    if (d.newOrders) persistNewOrders(d.newOrders);
    showToast('Restored from backup file');
  }

  function addDealer(d) {
    if (dealers.some(x => normName(x.name) === normName(d.name))) {
      showToast('Dealer already exists');
      return;
    }
    persistDealers([...dealers, d].sort((a, b) => a.name.localeCompare(b.name)));
    showToast(`Added dealer ${d.name}`);
    setDealerModal(false);
  }

  function toggleDealerOrigin(name) {
    const next = dealers.map(d => d.name === name ? { ...d, origin: d.origin === 'MX' ? 'US' : 'MX' } : d);
    persistDealers(next);
  }

  function renameDealer(oldName, newName) {
    if (!newName || newName === oldName) return;
    if (dealers.some(d => normName(d.name) === normName(newName) && d.name !== oldName)) {
      showToast('A dealer with that name already exists');
      return;
    }
    const nextDealers = dealers
      .map(d => d.name === oldName ? { ...d, name: newName } : d)
      .sort((a, b) => a.name.localeCompare(b.name));
    persistDealers(nextDealers);
    const nextOrders = orders.map(o => normName(o.dealer) === normName(oldName) ? { ...o, dealer: newName } : o);
    persistOrders(nextOrders);
    // Keep their portal login (and any active session) pointed at the new name — best effort,
    // only works on the deployed site, silently does nothing in the Claude preview.
    fetch('/api/dealer-auth', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'rename-dealer', oldName, newName, token: localStorage.getItem('gr-admin-token') || '' }),
    }).catch(() => {});
    showToast(`Renamed dealer to ${newName}`);
  }

  function deleteDealer(name) {
    persistDealers(dealers.filter(d => d.name !== name));
    // Revoke their portal login and log them out of any active session — best effort,
    // only works on the deployed site, silently does nothing in the Claude preview.
    fetch('/api/dealer-auth', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'remove-credentials', dealerName: name, token: localStorage.getItem('gr-admin-token') || '' }),
    }).catch(() => {});
    showToast(`Removed dealer ${name}`);
  }

  function addModel(sku) {
    const clean = sku.trim().toUpperCase();
    if (!clean) return;
    if (inventory.some(r => r.sku === clean)) {
      showToast('That model already exists');
      return;
    }
    const next = [...inventory, { sku: clean, mx: 0, us: 0 }].sort((a, b) => a.sku.localeCompare(b.sku));
    persistInventory(next);
    showToast(`Added model ${clean}`);
  }

  function renameModel(oldSku, newSku) {
    const clean = newSku.trim().toUpperCase();
    if (!clean || clean === oldSku) return;
    const oldRow = inventory.find(r => r.sku === oldSku);
    if (!oldRow) return;
    const existing = inventory.find(r => r.sku === clean);
    let nextInv;
    if (existing) {
      nextInv = inventory
        .filter(r => r.sku !== oldSku)
        .map(r => r.sku === clean ? { ...r, mx: (r.mx || 0) + (oldRow.mx || 0), us: (r.us || 0) + (oldRow.us || 0) } : r);
    } else {
      nextInv = inventory.map(r => r.sku === oldSku ? { ...r, sku: clean } : r);
    }
    persistInventory(nextInv);

    persistOrders(orders.map(o => o.sku === oldSku ? { ...o, sku: clean } : o));

    if (productionBatches.some(b => b.sku === oldSku)) {
      persistProductionBatches(productionBatches.map(b => b.sku === oldSku ? { ...b, sku: clean } : b));
    }

    persistSales(salesLog.map(s => s.sku === oldSku ? { ...s, sku: clean } : s));
    persistProductionLog(productionLog.map(p => p.sku === oldSku ? { ...p, sku: clean } : p));

    showToast(`Renamed ${oldSku} to ${clean}`);
  }

  function deleteModel(sku) {
    persistInventory(inventory.filter(r => r.sku !== sku));
    showToast(`Removed model ${sku}`);
  }

  function resetData(categories) {
    const set = new Set(categories);
    if (set.has('models')) { persistInventory([]); persistToolboxItems([]); }
    if (set.has('dealers')) { persistDealers([]); }
    if (set.has('orders')) { persistOrders([]); }
    if (set.has('sold')) { persistSales([]); }
    if (set.has('production')) { persistProductionBatches([]); persistProductionLog([]); }
    if (set.has('shipped')) { persistShippedLog([]); }
    showToast(`Deleted: ${categories.join(', ')}`);
  }

  if (loading) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: 400, fontFamily: "'Inter', sans-serif", color: '#5B6470', gap: 10 }}>
        <Loader2 className="spin" size={20} />
        Loading warehouse data…
        <style>{`.spin { animation: spin 1s linear infinite; } @keyframes spin { to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  const issueCount = dataIssues.unknownDealers.length + dataIssues.unknownSkus.length;
  const isOwner = !adminSession || adminSession.role === 'owner';
  const perms = isOwner
    ? { manageDealerLogins: true, processNewOrders: true, createOrders: true, editOrders: true, editModelsAndDealers: true, receiveOrders: true }
    : {
        manageDealerLogins: !!adminSession.permissions?.manageDealerLogins,
        processNewOrders: !!adminSession.permissions?.processNewOrders,
        createOrders: !!adminSession.permissions?.createOrders,
        editOrders: !!adminSession.permissions?.editOrders,
        editModelsAndDealers: !!adminSession.permissions?.editModelsAndDealers,
        receiveOrders: !!adminSession.permissions?.receiveOrders,
      };

  const NAV = [
    { group: 'Dealer requests', items: [
      { id: 'neworders', label: 'New Orders', icon: Inbox, count: newOrders.length },
    ] },
    { group: 'Look up', items: [
      { id: 'sku', label: 'Bumper Lookup', icon: PackageSearch },
      { id: 'dealer', label: 'Dealer Lookup', icon: UserSearch },
    ] },
    { group: 'Plan', items: [
      { id: 'production', label: 'Production Planning', icon: Factory },
      { id: 'sold', label: 'Sold Units', icon: TrendingDown },
      { id: 'orders', label: 'Orders', icon: ClipboardList },
      { id: 'shipped', label: 'Shipped', icon: PackageCheck },
      { id: 'models', label: 'Models', icon: Package },
      { id: 'dealers', label: 'Dealers', icon: Users },
    ] },
    { group: 'Data', items: [
      { id: 'import', label: 'Import', icon: Upload },
      { id: 'issues', label: 'Issues', icon: AlertTriangle, count: issueCount },
      { id: 'backup', label: 'Backup', icon: DownloadCloud },
      { id: 'reset', label: 'Reset Data', icon: Trash2 },
      ...(isOwner ? [{ id: 'staff', label: 'Staff', icon: Users }] : []),
    ] },
  ];

  return (
    <div style={{ fontFamily: "'Inter', sans-serif", background: '#EEEEE9', color: '#1C2126' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        input, select { font-family: 'Inter', sans-serif; }
        table { border-collapse: collapse; width: 100%; }
        .row-hover:hover { background: #F7F5EF !important; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-thumb { background: #C9C5B8; border-radius: 4px; }
        button { cursor: pointer; font-family: 'Inter', sans-serif; }
        .app-shell { display: flex; align-items: flex-start; min-height: 100%; width: 100%; }
        .sidebar { width: 220px; flex-shrink: 0; background: #1C2126; color: #F5F3EE; min-height: 100vh; padding: 20px 12px; position: sticky; top: 0; }
        .sidebar-nav-group-label { font-size: 10.5px; font-weight: 700; letter-spacing: 0.07em; text-transform: uppercase; color: #6B7280; padding: 14px 12px 6px; }
        .sidebar-nav-item { display: flex; align-items: center; gap: 10px; padding: 9px 12px; border-radius: 7px; font-size: 13px; font-weight: 600; color: #B7BCC2; cursor: pointer; margin-bottom: 2px; border: none; background: transparent; width: 100%; text-align: left; }
        .sidebar-nav-item:hover { background: #23282F; color: #F5F3EE; }
        .sidebar-nav-item.active { background: #E8592A; color: white; }
        .main-content { flex: 1; padding: 14px 18px; min-width: 0; }
        @media (max-width: 780px) {
          .app-shell { flex-direction: column; }
          .sidebar { width: 100%; min-height: auto; position: static; display: flex; flex-direction: row; align-items: center; overflow-x: auto; padding: 10px; gap: 4px; }
          .sidebar-brand { display: none; }
          .sidebar-nav-group-label { display: none; }
          .sidebar-nav-item { flex-shrink: 0; margin-bottom: 0; white-space: nowrap; }
          .main-content { padding: 10px; }
        }
      `}</style>

      <div className="app-shell">
        <div className="sidebar">
          <div className="sidebar-brand" style={{ padding: '10px 12px 16px' }}>
            <div style={{ background: 'white', borderRadius: 8, padding: 14 }}>
              <img src={LOGO_DATA_URL} alt="GR Trailers" style={{ width: '100%', height: 'auto', display: 'block', borderRadius: 3 }} />
            </div>
            <div style={{ fontSize: 10.5, color: '#6B7280', marginTop: 8 }}>MX &amp; GR/US</div>
          </div>
          {NAV.map(group => (
            <div key={group.group}>
              <div className="sidebar-nav-group-label">{group.group}</div>
              {group.items.map(t => (
                <button key={t.id} className={`sidebar-nav-item${tab === t.id ? ' active' : ''}`} onClick={() => setTab(t.id)}>
                  <t.icon size={15} />
                  <span style={{ flex: 1 }}>{t.label}</span>
                  {!!t.count && (
                    <span style={{
                      background: tab === t.id ? 'rgba(255,255,255,0.25)' : '#B23A2E', color: 'white',
                      borderRadius: 9, fontSize: 10, padding: '1px 6px', fontWeight: 700
                    }}>{t.count}</span>
                  )}
                </button>
              ))}
            </div>
          ))}
          {onAdminLogout && (
            <button onClick={onAdminLogout} className="sidebar-nav-item" style={{ marginTop: 10, borderTop: '1px solid #2A3038', paddingTop: 12 }}>
              <LogOut size={15} />
              <span style={{ flex: 1 }}>Log out</span>
            </button>
          )}
        </div>

        <div className="main-content">
          {tab === 'neworders' && (
            <NewOrdersView newOrders={newOrders} onProcess={processNewOrders} onReject={rejectNewOrders} canProcess={perms.processNewOrders} />
          )}
          {tab === 'sku' && (
            <SkuLookupView inventory={inventory} openOrders={openOrders} setShipModal={setShipModal} canShip={perms.receiveOrders} />
          )}
          {tab === 'dealer' && (
            <DealerLookupView dealers={dealers} openOrders={openOrders} inventory={inventory} ordersByLocSku={ordersByLocSku} setShipModal={setShipModal} canShip={perms.receiveOrders} />
          )}
          {tab === 'production' && (
            <ProductionPlanningView inventory={inventory} demandByLocSku={demandByLocSku} pendingBySku={pendingBySku} productionBatches={productionBatches} onSaveProduction={saveProductionOrder} onUpdateBatch={updateProductionBatch} onDeleteBatch={deleteProductionBatch} />
          )}
          {tab === 'sold' && (
            <SoldUnitsView salesLog={salesLog} />
          )}
          {tab === 'orders' && (
            <OrdersView orders={ordersEnriched} onAdd={() => setOrderModal(true)} setShipModal={setShipModal} setEditOrderModal={setEditOrderModal} canCreate={perms.createOrders} canEdit={perms.editOrders} canShip={perms.receiveOrders} />
          )}
          {tab === 'shipped' && (
            <ShippedView shippedLog={shippedLog} />
          )}
          {tab === 'models' && (
            <ModelsView inventory={inventory} orders={orders} pendingBySku={pendingBySku} onAdd={addModel} onRename={renameModel} onDelete={deleteModel} canEdit={perms.editModelsAndDealers} />
          )}
          {tab === 'dealers' && (
            <DealersView dealers={dealers} orders={orders} onAdd={() => setDealerModal(true)} toggleDealerOrigin={toggleDealerOrigin} onRename={renameDealer} onDelete={deleteDealer} onSetLogin={setDealerLoginModal} canManageDealerLogins={perms.manageDealerLogins} canEdit={perms.editModelsAndDealers} />
          )}
          {tab === 'import' && (
            <ImportView openImport={setImportModal} />
          )}
          {tab === 'issues' && (
            <IssuesView dataIssues={dataIssues} openOrders={openOrders} />
          )}
          {tab === 'backup' && (
            <BackupView state={{ inventory, toolboxItems, dealers, orders, salesLog, productionBatches, productionLog, newOrders }} onRestoreFile={restoreFromFile} />
          )}
          {tab === 'staff' && isOwner && (
            <StaffView />
          )}
          {tab === 'reset' && (
            <ResetDataView onReset={resetData} />
          )}
        </div>
      </div>

      {shipModal && (
        <ShipModal order={shipModal} onClose={() => setShipModal(null)} onShip={shipOrderQty} />
      )}
      {orderModal && (
        <OrderModal dealers={dealers} inventory={inventory} onClose={() => setOrderModal(false)} onAdd={addOrder} />
      )}
      {dealerModal && (
        <DealerModal onClose={() => setDealerModal(false)} onAdd={addDealer} />
      )}
      {(importModal === 'MX' || importModal === 'US') && (
        <ImportModal location={importModal} inventory={inventory} pendingBySku={pendingBySku} onClose={() => setImportModal(null)} onApply={applyImport} />
      )}
      {importModal === 'ORDERS' && (
        <OrdersImportModal onClose={() => setImportModal(null)} onApply={applyOrdersImport} />
      )}
      {dealerLoginModal && (
        <DealerLoginModal dealerName={dealerLoginModal} onClose={() => setDealerLoginModal(null)} />
      )}
      {editOrderModal && (
        <EditOrderModal order={editOrderModal} dealers={dealers} inventory={inventory} onClose={() => setEditOrderModal(null)} onSave={updateOrder} onDelete={deleteOrder} />
      )}

      {toast && <Toast message={toast} onDone={() => setToast(null)} />}
    </div>
  );
}

function computeCoverage(lineOrders, onHand) {
  const sorted = lineOrders.slice().sort((a, b) => {
    const da = parseDate(a.date)?.getTime() ?? Infinity;
    const db = parseDate(b.date)?.getTime() ?? Infinity;
    return da - db;
  });
  let running = onHand;
  return sorted.map(o => {
    const coveredQty = Math.max(0, Math.min(running, o.backordered));
    const fullyCovered = coveredQty === o.backordered;
    running -= coveredQty;
    return { ...o, coveredQty, fullyCovered, waitingQty: o.backordered - coveredQty };
  });
}

function th() { return { padding: '5px 9px', textAlign: 'left', fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.03em' }; }
function td() { return { padding: '5px 9px', fontSize: 12, verticalAlign: 'middle' }; }

function ProductionTable({ title, color, rows, orderQty, setQty, pendingBySku }) {
  if (rows.length === 0) return null;
  return (
    <div style={{ marginBottom: 10 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
        <div style={{ width: 8, height: 8, borderRadius: 2, background: color }} />
        <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 12.5, textTransform: 'uppercase', letterSpacing: '0.03em', color: '#5B6470' }}>{title}</div>
      </div>
      <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'auto' }}>
        <table>
          <thead>
            <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
              <th style={th()}>Model</th>
              <th style={{ ...th(), textAlign: 'right' }}>MX On Hand</th>
              <th style={{ ...th(), textAlign: 'right' }}>MX Demand</th>
              <th style={{ ...th(), textAlign: 'right' }}>MX Net</th>
              <th style={{ ...th(), textAlign: 'right', borderLeft: '1px solid #33404A' }}>US On Hand</th>
              <th style={{ ...th(), textAlign: 'right' }}>US Demand</th>
              <th style={{ ...th(), textAlign: 'right' }}>US Net</th>
              <th style={{ ...th(), textAlign: 'right', borderLeft: '1px solid #33404A' }}>Total Net</th>
              <th style={{ ...th(), textAlign: 'right', borderLeft: '1px solid #33404A' }}>Order Qty</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => {
              const pending = pendingBySku[r.sku] || 0;
              return (
                <tr key={r.sku} className="row-hover" style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                  <td style={td()}><SkuTag sku={r.sku} /></td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace" }}>
                    {r.mx}
                    {pending > 0 && <span title={`${pending} in production, not yet received`} style={{ fontSize: 10, color: '#B58A2E', marginLeft: 3 }}>+{pending}</span>}
                  </td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", color: r.mxDemand > 0 ? '#1C2126' : '#C9C5B8' }}>{r.mxDemand || '—'}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: r.mxNet < 0 ? '#B23A2E' : '#3E7B4F' }}>{r.mxNet > 0 ? '+' : ''}{r.mxNet}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", borderLeft: '1px solid #EFEDE4' }}>{r.us}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", color: r.usDemand > 0 ? '#1C2126' : '#C9C5B8' }}>{r.usDemand || '—'}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: r.usNet < 0 ? '#B23A2E' : '#3E7B4F' }}>{r.usNet > 0 ? '+' : ''}{r.usNet}</td>
                  <td style={{ ...td(), textAlign: 'right', borderLeft: '1px solid #EFEDE4' }}>
                    <span style={{
                      fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 13.5,
                      padding: '2px 8px', borderRadius: 5,
                      color: r.totalNet < 0 ? '#B23A2E' : '#3E7B4F',
                      background: r.totalNet < 0 ? '#FCEEE8' : '#EAF4EC'
                    }}>{r.totalNet > 0 ? '+' : ''}{r.totalNet}</span>
                  </td>
                  <td style={{ ...td(), textAlign: 'right' }}>
                    <input type="number" min="0" value={orderQty[r.sku] || ''} onChange={e => setQty(r.sku, e.target.value)}
                      placeholder="0" style={{
                        width: 52, textAlign: 'right', padding: '4px 6px', borderRadius: 5,
                        border: '1px solid #DCD9CE', fontSize: 12, fontFamily: "'IBM Plex Mono', monospace"
                      }} />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function InProductionPanel({ productionBatches, onUpdateBatch, onDeleteBatch }) {
  const [editMode, setEditMode] = useState(false);
  const batches = productionBatches
    .filter(b => b.qty > 0)
    .sort((a, b) => a.sku.localeCompare(b.sku) || ((parseDate(a.date)?.getTime() ?? -Infinity) - (parseDate(b.date)?.getTime() ?? -Infinity)));
  const total = batches.reduce((s, b) => s + b.qty, 0);
  const modelCount = new Set(batches.map(b => b.sku)).size;

  const columns = [];
  for (let i = 0; i < batches.length; i += 5) columns.push(batches.slice(i, i + 5));

  return (
    <div style={{ background: '#1C2126', borderRadius: 10, padding: '10px 14px', marginBottom: 12 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: (batches.length || editMode) ? 8 : 0, flexWrap: 'wrap', gap: 8 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <Factory size={13} color="#B7BCC2" />
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 12, textTransform: 'uppercase', letterSpacing: '0.03em', color: '#F5F3EE' }}>In Production</div>
          {total > 0 && <span style={{ fontSize: 11, color: '#8A8F97' }}>· {total} units across {modelCount} model{modelCount === 1 ? '' : 's'}</span>}
        </div>
        {batches.length > 0 && (
          <button onClick={() => setEditMode(e => !e)} style={{
            display: 'flex', alignItems: 'center', gap: 5,
            background: editMode ? '#E8592A' : '#2A3038', color: editMode ? 'white' : '#B7BCC2',
            border: 'none', borderRadius: 6, padding: '4px 10px', fontSize: 11, fontWeight: 700
          }}><Edit3 size={11} /> {editMode ? 'Done' : 'Edit'}</button>
        )}
      </div>

      {batches.length === 0 ? (
        <div style={{ fontSize: 12, color: '#8A8F97' }}>Nothing pending — save a pedido below to send models to production.</div>
      ) : editMode ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          {batches.map(b => (
            <div key={b.id} style={{ display: 'flex', alignItems: 'center', gap: 8, background: '#23282F', borderRadius: 7, padding: '5px 8px' }}>
              <div style={{ width: 130 }}><SkuTag sku={b.sku} /></div>
              <input type="number" min="0" value={b.qty} onChange={e => onUpdateBatch(b.id, { qty: Math.max(0, parseInt(e.target.value, 10) || 0) })} style={{
                width: 60, padding: '4px 6px', borderRadius: 5, border: '1px solid #3A414B', background: '#1C2126', color: '#F5F3EE',
                fontFamily: "'IBM Plex Mono', monospace", fontSize: 12
              }} />
              <input type="date" value={usToISO(b.date)} onChange={e => onUpdateBatch(b.id, { date: isoToUS(e.target.value) })} style={{
                padding: '4px 6px', borderRadius: 5, border: '1px solid #3A414B', background: '#1C2126', color: '#F5F3EE', fontSize: 11.5
              }} />
              <button onClick={() => onDeleteBatch(b.id)} style={{
                marginLeft: 'auto', background: '#3A2020', border: '1px solid #5C2C2C', color: '#E08080',
                borderRadius: 5, padding: '4px 7px', display: 'flex', alignItems: 'center'
              }}><X size={12} /></button>
            </div>
          ))}
        </div>
      ) : (
        <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap' }}>
          {columns.map((col, i) => (
            <div key={i} style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, color: '#F5F3EE', lineHeight: 1.7 }}>
              {col.map(b => (
                <div key={b.id}>
                  {b.sku} = {b.qty}
                  {b.date && <span style={{ color: '#8A8F97', fontSize: 11 }}> ({b.date})</span>}
                </div>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function ProductionPlanningView({ inventory, demandByLocSku, pendingBySku, productionBatches, onSaveProduction, onUpdateBatch, onDeleteBatch }) {
  const [search, setSearch] = useState('');
  const [sortMode, setSortMode] = useState('urgent');
  const [orderQty, setOrderQty] = useState({});
  const [orderDate, setOrderDate] = useState(todayISO());

  const setQty = (sku, val) => setOrderQty(prev => ({ ...prev, [sku]: val }));

  const rows = useMemo(() => {
    const list = inventory
      .filter(r => r.sku.toLowerCase().includes(search.toLowerCase()))
      .map(r => {
        const mxDemand = demandByLocSku['MX|' + r.sku] || 0;
        const usDemand = demandByLocSku['US|' + r.sku] || 0;
        const mxNet = (r.mx || 0) - mxDemand;
        const usNet = (r.us || 0) - usDemand;
        return { ...r, mxDemand, usDemand, mxNet, usNet, totalNet: mxNet + usNet };
      });
    if (sortMode === 'urgent') list.sort((a, b) => a.totalNet - b.totalNet);
    else list.sort((a, b) => a.sku.localeCompare(b.sku));
    return list;
  }, [inventory, demandByLocSku, search, sortMode]);

  const fbRows = rows.filter(r => r.sku.toUpperCase().startsWith('FB'));
  const rbRows = rows.filter(r => r.sku.toUpperCase().startsWith('RB'));
  const otherRows = rows.filter(r => !r.sku.toUpperCase().startsWith('FB') && !r.sku.toUpperCase().startsWith('RB'));

  const toProduce = rows.filter(r => r.totalNet < 0).length;

  const pedidoLines = Object.entries(orderQty)
    .filter(([sku, val]) => val && parseInt(val, 10) > 0 && inventory.some(r => r.sku === sku))
    .sort((a, b) => a[0].localeCompare(b[0]));

  function handleSave() {
    const items = pedidoLines.map(([sku, val]) => ({ sku, qty: parseInt(val, 10) }));
    onSaveProduction(items, isoToUS(orderDate));
    setOrderQty({});
    setOrderDate(todayISO());
  }

  return (
    <div>
      <div style={{ marginBottom: 6 }}>
        <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Production Planning</div>
        <div style={{ fontSize: 12.5, color: '#5B6470' }}>
          Current inventory minus open orders, both warehouses. <strong style={{ color: toProduce > 0 ? '#B23A2E' : '#3E7B4F' }}>{toProduce}</strong> model{toProduce === 1 ? '' : 's'} running a deficit across MX + US.
        </div>
      </div>

      <InProductionPanel productionBatches={productionBatches} onUpdateBatch={onUpdateBatch} onDeleteBatch={onDeleteBatch} />

      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 7 }}>
        <div style={{ position: 'relative', flex: 1, maxWidth: 320 }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: '#8A8F97' }} />
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search model…" style={{
            width: '100%', padding: '6px 28px 6px 26px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, background: 'white'
          }} />
          {search && <ClearButton onClick={() => setSearch('')} />}
        </div>
        <select value={sortMode} onChange={e => setSortMode(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13 }}>
          <option value="urgent">Sort: most urgent first</option>
          <option value="sku">Sort: Model A–Z</option>
        </select>
      </div>

      <ProductionTable title="FB Models" color="#33546E" rows={fbRows} orderQty={orderQty} setQty={setQty} pendingBySku={pendingBySku} />
      <ProductionTable title="RB Models" color="#B23A2E" rows={rbRows} orderQty={orderQty} setQty={setQty} pendingBySku={pendingBySku} />
      <ProductionTable title="Other" color="#8A8F97" rows={otherRows} orderQty={orderQty} setQty={setQty} pendingBySku={pendingBySku} />

      <div style={{ marginTop: 20, background: '#1C2126', borderRadius: 10, padding: '14px 16px' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: pedidoLines.length ? 8 : 0, flexWrap: 'wrap', gap: 8 }}>
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 14, textTransform: 'uppercase', color: '#F5F3EE' }}>
            Pedido de defensas
          </div>
          {pedidoLines.length > 0 && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11.5, color: '#B7BCC2' }}>
                Date
                <input type="date" value={orderDate} onChange={e => setOrderDate(e.target.value)} style={{
                  padding: '4px 6px', borderRadius: 5, border: '1px solid #3A414B', background: '#2A3038', color: '#F5F3EE', fontSize: 12
                }} />
              </label>
              <button onClick={handleSave} style={{
                display: 'flex', alignItems: 'center', gap: 6, background: '#E8592A', color: 'white',
                border: 'none', borderRadius: 7, padding: '7px 12px', fontSize: 12, fontWeight: 700
              }}><Check size={13} /> Save</button>
            </div>
          )}
        </div>
        {pedidoLines.length === 0 ? (
          <div style={{ fontSize: 12, color: '#8A8F97' }}>Type a quantity in Order Qty above to add it here.</div>
        ) : (
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: '#F5F3EE', lineHeight: 1.9 }}>
            {pedidoLines.map(([sku, val]) => (
              <div key={sku}>{sku} = {parseInt(val, 10)}</div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function SkuLookupView({ inventory, openOrders, setShipModal, canShip }) {
  const [selected, setSelected] = useState('');
  const options = useMemo(() => inventory.map(r => r.sku).sort(), [inventory]);
  const row = inventory.find(r => r.sku === selected);

  const combined = useMemo(() => {
    if (!row) return [];
    const forSku = openOrders.filter(o => o.sku === selected);
    const mxOrders = forSku.filter(o => o.shipFrom === 'MX');
    const usOrders = forSku.filter(o => o.shipFrom === 'US');
    const mxCov = computeCoverage(mxOrders, row.mx || 0);
    const usCov = computeCoverage(usOrders, row.us || 0);
    return [...mxCov, ...usCov].sort((a, b) => (parseDate(a.date)?.getTime() ?? Infinity) - (parseDate(b.date)?.getTime() ?? Infinity));
  }, [row, openOrders, selected]);

  const totalOwed = combined.reduce((s, o) => s + o.backordered, 0);

  return (
    <div>
      <div style={{ marginBottom: 6, maxWidth: 380 }}>
        <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 6 }}>Bumper Lookup</div>
        <label style={labelStyle()}>Look up a bumper model</label>
        <div style={{ position: 'relative' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 12, color: '#8A8F97' }} />
          <input list="sku-lookup-list" value={selected} onChange={e => setSelected(e.target.value)}
            placeholder="Start typing a model…"
            style={{ ...fieldStyle(), paddingLeft: 30, paddingRight: 28, marginBottom: 0 }} />
          {selected && <ClearButton onClick={() => setSelected('')} />}
        </div>
        <datalist id="sku-lookup-list">
          {options.map(s => <option key={s} value={s} />)}
        </datalist>
      </div>

      {selected && !row && (
        <div style={{ color: '#8A8F97', fontSize: 13 }}>No model "{selected}" found — pick one from the list.</div>
      )}

      {row && (
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 8, flexWrap: 'wrap' }}>
            <SkuTag sku={row.sku} />
            <div style={{ display: 'flex', gap: 18 }}>
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: LOCATIONS.MX.color }}>{row.mx}</div>
                <div style={{ fontSize: 11, color: '#5B6470' }}>on hand · MX</div>
              </div>
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace", color: LOCATIONS.US.color }}>{row.us}</div>
                <div style={{ fontSize: 11, color: '#5B6470' }}>on hand · US</div>
              </div>
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace" }}>{totalOwed}</div>
                <div style={{ fontSize: 11, color: '#5B6470' }}>units owed</div>
              </div>
            </div>
          </div>

          {combined.length === 0 ? (
            <div style={{ color: '#5B6470', fontSize: 13, background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 14 }}>
              No open orders for this model right now.
            </div>
          ) : (
            <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'hidden' }}>
              <table>
                <thead>
                  <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
                    <th style={th()}>Dealer</th>
                    <th style={th()}>PO</th>
                    <th style={th()}>Ordered</th>
                    <th style={{ ...th(), textAlign: 'right' }}>Days waiting</th>
                    <th style={th()}>Ships from</th>
                    <th style={{ ...th(), textAlign: 'right' }}>Owed</th>
                    <th style={th()}>Status</th>
                    <th style={th()}></th>
                  </tr>
                </thead>
                <tbody>
                  {combined.map((o, i) => {
                    const days = daysSince(o.date);
                    const statusColor = o.fullyCovered ? '#3E7B4F' : (o.coveredQty > 0 ? '#B58A2E' : '#B23A2E');
                    const statusLabel = o.fullyCovered ? 'Ready to ship' : (o.coveredQty > 0 ? `Partial (${o.coveredQty}/${o.backordered})` : 'Waiting on stock');
                    return (
                      <tr key={o.id} className="row-hover" style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                        <td style={td()}>
                          {o.dealer}
                          {!o.dealerKnown && <AlertTriangle size={11} color="#B23A2E" style={{ marginLeft: 6, display: 'inline' }} />}
                        </td>
                        <td style={{ ...td(), fontSize: 12.5, color: '#8A8F97' }}>{o.po || '—'}</td>
                        <td style={{ ...td(), fontSize: 12.5, color: '#5B6470' }}>{o.date || '—'}</td>
                        <td style={{ ...td(), textAlign: 'right', fontSize: 12.5, fontWeight: days > 30 ? 700 : 400, color: days > 30 ? '#B23A2E' : '#5B6470' }}>{days ?? '—'}</td>
                        <td style={td()}><Badge color={LOCATIONS[o.shipFrom].color} bg="transparent" border={LOCATIONS[o.shipFrom].color}>{o.shipFrom}</Badge></td>
                        <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{o.backordered}</td>
                        <td style={td()}><Badge color={statusColor} bg="transparent" border={statusColor}>{statusLabel}</Badge></td>
                        <td style={td()}>
                          {canShip && (
                            <button onClick={() => setShipModal(o)} style={{
                              fontSize: 11, fontWeight: 700, color: '#33546E', background: '#EAF0F4', border: '1px solid #C7D6DE',
                              borderRadius: 5, padding: '3px 8px', display: 'flex', alignItems: 'center', gap: 4
                            }}><ArrowRight size={11} /> Ship</button>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function OrdersView({ orders, onAdd, setShipModal, setEditOrderModal, canCreate, canEdit, canShip }) {
  const [filter, setFilter] = useState('open');
  const [q, setQ] = useState('');
  const rows = orders
    .filter(o => filter === 'all' ? true : o.backordered > 0)
    .filter(o => (o.dealer + o.sku + o.po + o.num).toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => b.id.localeCompare(a.id));

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <div style={{ position: 'relative' }}>
            <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: '#8A8F97' }} />
            <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search dealer, model, PO…" style={{
              padding: '6px 28px 6px 26px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, width: 260
            }} />
            {q && <ClearButton onClick={() => setQ('')} />}
          </div>
          <select value={filter} onChange={e => setFilter(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13 }}>
            <option value="open">Open only</option>
            <option value="all">All orders</option>
          </select>
        </div>
        {canCreate && (
          <button onClick={onAdd} style={{
            display: 'flex', alignItems: 'center', gap: 6, background: '#E8592A', color: 'white',
            border: 'none', borderRadius: 7, padding: '7px 12px', fontSize: 13, fontWeight: 700
          }}><Plus size={14} /> New Order</button>
        )}
      </div>

      <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'hidden' }}>
        <table>
          <thead>
            <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
              <th style={th()}>Model</th>
              <th style={th()}>Dealer</th>
              <th style={th()}>PO</th>
              <th style={th()}>Ships from</th>
              <th style={th()}>Date</th>
              <th style={{ ...th(), textAlign: 'right' }}>Qty</th>
              <th style={{ ...th(), textAlign: 'right' }}>Shipped</th>
              <th style={{ ...th(), textAlign: 'right' }}>Owed</th>
              <th style={th()}></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((o, i) => (
              <tr key={o.id} className="row-hover" style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                <td style={td()}><SkuTag sku={o.sku} />{!o.skuKnown && <AlertTriangle size={11} color="#B23A2E" style={{ marginLeft: 6, display: 'inline' }} />}</td>
                <td style={td()}>{o.dealer}{!o.dealerKnown && <AlertTriangle size={11} color="#B23A2E" style={{ marginLeft: 6, display: 'inline' }} />}</td>
                <td style={{ ...td(), color: '#8A8F97' }}>{o.po || '—'}</td>
                <td style={td()}><Badge color={LOCATIONS[o.shipFrom].color} bg="transparent" border={LOCATIONS[o.shipFrom].color}>{o.shipFrom}</Badge></td>
                <td style={{ ...td(), fontSize: 12.5, color: '#5B6470' }}>{o.date || '—'}</td>
                <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace" }}>{o.qty}</td>
                <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", color: '#3E7B4F' }}>{o.invoiced}</td>
                <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: o.backordered > 0 ? '#B23A2E' : '#C9C5B8' }}>{o.backordered}</td>
                <td style={td()}>
                  <span style={{ display: 'flex', gap: 6, justifyContent: 'flex-start' }}>
                    {o.backordered > 0 && canShip && (
                      <button onClick={() => setShipModal(o)} style={{
                        fontSize: 11, fontWeight: 700, color: '#33546E', background: '#EAF0F4', border: '1px solid #C7D6DE',
                        borderRadius: 5, padding: '3px 8px', display: 'flex', alignItems: 'center', gap: 4
                      }}><ArrowRight size={11} /> Ship</button>
                    )}
                    {canEdit && (
                      <button onClick={() => setEditOrderModal(o)} style={{
                        fontSize: 11, fontWeight: 700, color: '#5B6470', background: '#F3F2EE', border: '1px solid #DCD9CE',
                        borderRadius: 5, padding: '3px 7px', display: 'flex', alignItems: 'center'
                      }}><Edit3 size={11} /></button>
                    )}
                  </span>
                </td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={8} style={{ ...td(), textAlign: 'center', color: '#8A8F97', padding: 30 }}>No orders match.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function DealersView({ dealers, orders, onAdd, toggleDealerOrigin, onRename, onDelete, onSetLogin, canManageDealerLogins, canEdit }) {
  const [q, setQ] = useState('');
  const [editMode, setEditMode] = useState(false);
  const [editingName, setEditingName] = useState(null);
  const [renameValue, setRenameValue] = useState('');
  const [confirmDeleteName, setConfirmDeleteName] = useState(null);
  const rows = dealers.filter(d => d.name.toLowerCase().includes(q.toLowerCase()))
    .sort((a, b) => a.name.localeCompare(b.name));

  function openOrderCount(name) {
    return orders.filter(o => normName(o.dealer) === normName(name) && o.backordered > 0).length;
  }
  function saveRename(oldName) {
    if (renameValue.trim() && renameValue.trim() !== oldName) onRename(oldName, renameValue.trim());
    setEditingName(null);
  }

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, gap: 10, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: '#8A8F97' }} />
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search dealer…" style={{
            padding: '6px 28px 6px 26px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, width: 260
          }} />
          {q && <ClearButton onClick={() => setQ('')} />}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ fontSize: 12, color: '#8A8F97' }}>{rows.length} dealers</div>
          {canEdit && (
            <>
              <button onClick={() => { setEditMode(e => !e); setEditingName(null); setConfirmDeleteName(null); }} style={{
                display: 'flex', alignItems: 'center', gap: 6,
                background: editMode ? '#1C2126' : 'white', color: editMode ? 'white' : '#1C2126',
                border: '1px solid #1C2126', borderRadius: 7, padding: '7px 12px', fontSize: 12.5, fontWeight: 700
              }}><Edit3 size={13} /> {editMode ? 'Done editing' : 'Edit dealers'}</button>
              <button onClick={onAdd} style={{
                display: 'flex', alignItems: 'center', gap: 6, background: '#E8592A', color: 'white',
                border: 'none', borderRadius: 7, padding: '7px 12px', fontSize: 13, fontWeight: 700
              }}><Plus size={14} /> New Dealer</button>
            </>
          )}
        </div>
      </div>
      {editMode && canEdit && (
        <div style={{ fontSize: 12, color: '#B23A2E', background: '#FCEEE8', border: '1px solid #F0C4B8', borderRadius: 7, padding: '8px 12px', marginBottom: 7 }}>
          Editing on — click a badge to flip shipping origin, the key icon to set a portal login, the pencil to rename, or the trash icon to delete.
        </div>
      )}
      <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'hidden', display: 'grid', gridTemplateColumns: editMode ? '1fr' : 'repeat(auto-fill, minmax(280px, 1fr))' }}>
        {rows.map((d, i) => {
          const isEditing = editingName === d.name;
          const isConfirmingDelete = confirmDeleteName === d.name;
          return (
            <div key={d.name} style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '7px 12px', gap: 10,
              borderTop: '1px solid #EFEDE4', borderLeft: (!editMode && i % 2 === 1) ? '1px solid #EFEDE4' : 'none'
            }}>
              {isEditing ? (
                <input autoFocus value={renameValue} onChange={e => setRenameValue(e.target.value)}
                  onKeyDown={e => { if (e.key === 'Enter') saveRename(d.name); if (e.key === 'Escape') setEditingName(null); }}
                  style={{ flex: 1, padding: '4px 8px', borderRadius: 5, border: '1px solid #33546E', fontSize: 13 }} />
              ) : (
                <span style={{ fontSize: 13 }}>{d.name}</span>
              )}

              {isConfirmingDelete ? (
                <span style={{ display: 'flex', gap: 6, alignItems: 'center', fontSize: 11.5, flexShrink: 0 }}>
                  <span style={{ color: '#B23A2E' }}>
                    {openOrderCount(d.name) > 0 ? `${openOrderCount(d.name)} open order(s) — delete anyway?` : 'Delete this dealer?'}
                  </span>
                  <button onClick={() => { onDelete(d.name); setConfirmDeleteName(null); }} style={{ background: '#B23A2E', color: 'white', border: 'none', borderRadius: 5, padding: '3px 8px', fontSize: 11, fontWeight: 700 }}>Yes</button>
                  <button onClick={() => setConfirmDeleteName(null)} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 5, padding: '3px 8px', fontSize: 11 }}>Cancel</button>
                </span>
              ) : isEditing ? (
                <span style={{ display: 'flex', gap: 6, flexShrink: 0 }}>
                  <button onClick={() => saveRename(d.name)} style={{ background: '#3E7B4F', color: 'white', border: 'none', borderRadius: 5, padding: '3px 8px', fontSize: 11, fontWeight: 700 }}>Save</button>
                  <button onClick={() => setEditingName(null)} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 5, padding: '3px 8px', fontSize: 11 }}>Cancel</button>
                </span>
              ) : editMode ? (
                <span style={{ display: 'flex', gap: 6, alignItems: 'center', flexShrink: 0 }}>
                  <button onClick={() => toggleDealerOrigin(d.name)} style={{ border: 'none', background: 'none', padding: 0, cursor: 'pointer' }}>
                    <Badge color={LOCATIONS[d.origin].color} bg="transparent" border={LOCATIONS[d.origin].color}>{d.origin}</Badge>
                  </button>
                  {canManageDealerLogins && (
                    <button onClick={() => onSetLogin(d.name)} title="Set portal login" style={{ background: '#F3EFE6', border: '1px solid #DCD2B8', color: '#8A6D1F', borderRadius: 5, padding: '3px 6px', display: 'flex', alignItems: 'center' }}><KeyRound size={11} /></button>
                  )}
                  <button onClick={() => { setEditingName(d.name); setRenameValue(d.name); }} style={{ background: '#EAF0F4', border: '1px solid #C7D6DE', color: '#33546E', borderRadius: 5, padding: '3px 6px', display: 'flex', alignItems: 'center' }}><Edit3 size={11} /></button>
                  <button onClick={() => setConfirmDeleteName(d.name)} style={{ background: '#FCEEE8', border: '1px solid #F0C4B8', color: '#B23A2E', borderRadius: 5, padding: '3px 6px', display: 'flex', alignItems: 'center' }}><X size={11} /></button>
                </span>
              ) : (
                <Badge color={LOCATIONS[d.origin].color} bg="transparent" border={LOCATIONS[d.origin].color}>{d.origin}</Badge>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function IssuesView({ dataIssues, openOrders }) {
  const { unknownDealers, unknownSkus } = dataIssues;
  if (unknownDealers.length === 0 && unknownSkus.length === 0) {
    return (
      <div style={{ textAlign: 'center', padding: 60, color: '#5B6470' }}>
        <PackageCheck size={32} color="#3E7B4F" style={{ marginBottom: 6 }} />
        <div style={{ fontSize: 14 }}>No data issues. Every open order matches a known dealer and model.</div>
      </div>
    );
  }
  return (
    <div style={{ display: 'grid', gap: 18 }}>
      {unknownDealers.length > 0 && (
        <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
            <AlertTriangle size={15} color="#B23A2E" />
            <span style={{ fontWeight: 700, fontSize: 13.5 }}>Dealers on orders that aren't in your Dealer list</span>
          </div>
          <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 6 }}>
            These orders default to shipping from Mexico until you add the dealer with the correct origin.
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
            {unknownDealers.map(n => <Badge key={n} color="#B23A2E" bg="#FCEEE8" border="#F0C4B8">{n}</Badge>)}
          </div>
        </div>
      )}
      {unknownSkus.length > 0 && (
        <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
            <PackageX size={15} color="#B23A2E" />
            <span style={{ fontWeight: 700, fontSize: 13.5 }}>Models on orders that don't exist in inventory</span>
          </div>
          <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 6 }}>
            Likely a naming mismatch between your order codes and warehouse codes — check the Translation sheet or rename to match.
          </div>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
            {unknownSkus.map(s => <SkuTag key={s} sku={s} />)}
          </div>
        </div>
      )}
    </div>
  );
}

function DealerLookupView({ dealers, openOrders, inventory, ordersByLocSku, setShipModal, canShip }) {
  const [selected, setSelected] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const options = useMemo(() => dealers.map(d => d.name).sort(), [dealers]);
  const dealerInfo = dealers.find(d => normName(d.name) === normName(selected));
  const allRows = useMemo(() => {
    if (!dealerInfo) return [];
    return openOrders
      .filter(o => normName(o.dealer) === normName(selected))
      .sort((a, b) => (parseDate(a.date)?.getTime() ?? Infinity) - (parseDate(b.date)?.getTime() ?? Infinity));
  }, [openOrders, selected, dealerInfo]);

  // Same priority logic as Bumper Lookup / Production Planning: for each model + warehouse,
  // stock gets allocated oldest-order-first across ALL dealers, not just this one — so this
  // dealer's spot in line depends on who else is waiting on the same model.
  const coverageById = useMemo(() => {
    const map = {};
    Object.entries(ordersByLocSku).forEach(([key, list]) => {
      const [loc, sku] = key.split('|');
      const invRow = inventory.find(r => r.sku === sku);
      const onHand = invRow ? (loc === 'MX' ? invRow.mx : invRow.us) : 0;
      computeCoverage(list, onHand).forEach(o => { map[o.id] = o; });
    });
    return map;
  }, [ordersByLocSku, inventory]);

  function statusKeyFor(o) {
    const cov = coverageById[o.id];
    if (!cov) return 'unknown';
    if (cov.fullyCovered) return 'ready';
    if (cov.coveredQty > 0) return 'partial';
    return 'waiting';
  }

  // Filtering never disturbs the oldest-to-newest order — it just narrows the same list down.
  const rows = statusFilter === 'all' ? allRows : allRows.filter(o => statusKeyFor(o) === statusFilter);
  const totalUnits = rows.reduce((s, o) => s + o.backordered, 0);

  return (
    <div>
      <div style={{ marginBottom: 6, maxWidth: 380 }}>
        <label style={labelStyle()}>Look up a dealer</label>
        <div style={{ position: 'relative' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 12, color: '#8A8F97' }} />
          <input list="dealer-lookup-list" value={selected} onChange={e => setSelected(e.target.value)}
            placeholder="Start typing a dealer name…"
            style={{ ...fieldStyle(), paddingLeft: 30, paddingRight: 28, marginBottom: 0 }} />
          {selected && <ClearButton onClick={() => setSelected('')} />}
        </div>
        <datalist id="dealer-lookup-list">
          {options.map(n => <option key={n} value={n} />)}
        </datalist>
      </div>

      {selected && !dealerInfo && (
        <div style={{ color: '#8A8F97', fontSize: 13 }}>No dealer named "{selected}" found — pick one from the list.</div>
      )}

      {dealerInfo && (
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
            <span style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 19, textTransform: 'uppercase' }}>{dealerInfo.name}</span>
            <Badge color={LOCATIONS[dealerInfo.origin].color} bg="transparent" border={LOCATIONS[dealerInfo.origin].color}>Ships from {LOCATIONS[dealerInfo.origin].short}</Badge>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, marginBottom: 8, flexWrap: 'wrap' }}>
            <div style={{ fontSize: 12.5, color: '#5B6470' }}>
              {rows.length} open order line{rows.length === 1 ? '' : 's'} · {totalUnits} unit{totalUnits === 1 ? '' : 's'} owed
              {statusFilter !== 'all' && allRows.length !== rows.length && <span> (of {allRows.length} total)</span>}
            </div>
            <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)} style={{
              padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 12.5
            }}>
              <option value="all">All statuses</option>
              <option value="ready">Ready to ship</option>
              <option value="partial">Partial</option>
              <option value="waiting">Waiting on stock</option>
            </select>
          </div>

          {rows.length === 0 ? (
            <div style={{ color: '#5B6470', fontSize: 13, background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 14 }}>
              {allRows.length === 0 ? 'No open orders for this dealer right now.' : 'No orders match this filter.'}
            </div>
          ) : (
            <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'hidden' }}>
              <table>
                <thead>
                  <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
                    <th style={th()}>Model</th>
                    <th style={th()}>Ordered</th>
                    <th style={th()}>PO</th>
                    <th style={{ ...th(), textAlign: 'right' }}>Owed</th>
                    <th style={{ ...th(), userSelect: 'none' }}>Ships from</th>
                    <th style={{ ...th(), textAlign: 'right', userSelect: 'none' }}>Days waiting</th>
                    <th style={{ ...th(), userSelect: 'none' }}>Status</th>
                    <th style={th()}></th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((o, i) => {
                    const days = daysSince(o.date);
                    const cov = coverageById[o.id];
                    const statusColor = cov ? (cov.fullyCovered ? '#3E7B4F' : (cov.coveredQty > 0 ? '#B58A2E' : '#B23A2E')) : '#8A8F97';
                    const statusLabel = cov
                      ? (cov.fullyCovered ? 'Ready to ship' : (cov.coveredQty > 0 ? `Partial (${cov.coveredQty}/${cov.backordered})` : 'Waiting on stock'))
                      : '—';
                    return (
                      <tr key={o.id} className="row-hover" style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                        <td style={td()}><SkuTag sku={o.sku} />{!o.skuKnown && <AlertTriangle size={11} color="#B23A2E" style={{ marginLeft: 6, display: 'inline' }} />}</td>
                        <td style={{ ...td(), fontSize: 12.5, color: '#5B6470' }}>{o.date || '—'}</td>
                        <td style={{ ...td(), fontSize: 12.5, color: '#8A8F97' }}>{o.po || '—'}</td>
                        <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700 }}>{o.backordered}</td>
                        <td style={{ ...td(), userSelect: 'none' }}><Badge color={LOCATIONS[o.shipFrom].color} bg="transparent" border={LOCATIONS[o.shipFrom].color}>{o.shipFrom}</Badge></td>
                        <td style={{ ...td(), textAlign: 'right', fontSize: 12.5, fontWeight: days > 30 ? 700 : 400, color: days > 30 ? '#B23A2E' : '#5B6470', userSelect: 'none' }}>{days ?? '—'}</td>
                        <td style={{ ...td(), userSelect: 'none' }}><Badge color={statusColor} bg="transparent" border={statusColor}>{statusLabel}</Badge></td>
                        <td style={td()}>
                          {canShip && (
                            <button onClick={() => setShipModal(o)} style={{
                              fontSize: 11, fontWeight: 700, color: '#33546E', background: '#EAF0F4', border: '1px solid #C7D6DE',
                              borderRadius: 5, padding: '3px 8px', display: 'flex', alignItems: 'center', gap: 4
                            }}><ArrowRight size={11} /> Ship</button>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function ImportView({ openImport }) {
  const cards = [
    {
      id: 'MX', title: 'Mexico Inventory', desc: 'Upload the "Existencia y valor del inventario" export. Matches Spanish item names to model codes via the Translation table and replaces MX on-hand counts.',
      color: LOCATIONS.MX.color,
    },
    {
      id: 'US', title: 'GR / US Inventory', desc: "Upload the GR/US stock export. Replaces US on-hand counts with what's in the file.",
      color: LOCATIONS.US.color,
    },
    {
      id: 'ORDERS', title: 'Orders', desc: "Upload the QuickBooks open sales orders export. Replaces the full order list — dealer, PO, dates, quantities, and what's still owed.",
      color: '#5B6470',
    },
  ];
  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Import</div>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 6 }}>
        Bring in fresh data from Excel. Each import shows a preview before anything is applied.
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: 14 }}>
        {cards.map(c => (
          <div key={c.id} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 12, padding: 13 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
              <div style={{ width: 32, height: 32, borderRadius: 7, background: c.color + '1A', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <FileSpreadsheet size={16} color={c.color} />
              </div>
              <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 14.5, textTransform: 'uppercase' }}>{c.title}</div>
            </div>
            <div style={{ fontSize: 12, color: '#5B6470', lineHeight: 1.5, marginBottom: 8, minHeight: 34 }}>{c.desc}</div>
            <button onClick={() => openImport(c.id)} style={{
              width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6,
              background: '#1C2126', color: 'white', border: 'none', borderRadius: 7, padding: '7px 12px', fontSize: 12.5, fontWeight: 700
            }}><Upload size={13} /> Import {c.title}</button>
          </div>
        ))}
      </div>
    </div>
  );
}

function SoldUnitsView({ salesLog }) {
  const [skuFilter, setSkuFilter] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [sortMode, setSortMode] = useState('date');

  const rows = useMemo(() => {
    const filtered = salesLog
      .filter(s => !skuFilter || s.sku.toLowerCase().includes(skuFilter.toLowerCase()))
      .filter(s => {
        const d = parseDate(s.date);
        if (!d) return true;
        if (from && d < new Date(from + 'T00:00:00')) return false;
        if (to && d > new Date(to + 'T23:59:59')) return false;
        return true;
      });
    if (sortMode === 'most') return filtered.sort((a, b) => b.qty - a.qty);
    if (sortMode === 'least') return filtered.sort((a, b) => a.qty - b.qty);
    return filtered.sort((a, b) => (parseDate(b.date)?.getTime() ?? 0) - (parseDate(a.date)?.getTime() ?? 0));
  }, [salesLog, skuFilter, from, to, sortMode]);

  const totalSold = rows.reduce((s, r) => s + r.qty, 0);

  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Sold Units</div>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 8 }}>
        Logged automatically whenever an inventory import shows fewer units than last time. <strong>{totalSold}</strong> unit{totalSold === 1 ? '' : 's'} sold in the selected range.
      </div>

      <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 8, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: '#8A8F97' }} />
          <input value={skuFilter} onChange={e => setSkuFilter(e.target.value)} placeholder="Filter by Model…" style={{
            padding: '6px 28px 6px 26px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, width: 200
          }} />
          {skuFilter && <ClearButton onClick={() => setSkuFilter('')} />}
        </div>
        <label style={{ fontSize: 12, color: '#5B6470', display: 'flex', alignItems: 'center', gap: 5 }}>
          From
          <input type="date" value={from} onChange={e => setFrom(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 12.5 }} />
        </label>
        <label style={{ fontSize: 12, color: '#5B6470', display: 'flex', alignItems: 'center', gap: 5 }}>
          To
          <input type="date" value={to} onChange={e => setTo(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 12.5 }} />
        </label>
        <select value={sortMode} onChange={e => setSortMode(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 12.5 }}>
          <option value="date">Sort: newest first</option>
          <option value="most">Sort: most sold</option>
          <option value="least">Sort: least sold</option>
        </select>
        {(from || to || skuFilter) && (
          <button onClick={() => { setFrom(''); setTo(''); setSkuFilter(''); }} style={{
            fontSize: 12, color: '#B23A2E', background: 'none', border: 'none', fontWeight: 600
          }}>Clear filters</button>
        )}
      </div>

      <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'hidden' }}>
        <table>
          <thead>
            <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
              <th style={th()}>Date</th>
              <th style={th()}>Model</th>
              <th style={th()}>Location</th>
              <th style={{ ...th(), textAlign: 'right' }}>Sold</th>
              <th style={{ ...th(), textAlign: 'right' }}>Before → After</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((s, i) => (
              <tr key={s.id} className="row-hover" style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                <td style={{ ...td(), fontSize: 12, color: '#5B6470' }}>{s.date}</td>
                <td style={td()}><SkuTag sku={s.sku} /></td>
                <td style={td()}><Badge color={LOCATIONS[s.location].color} bg="transparent" border={LOCATIONS[s.location].color}>{s.location}</Badge></td>
                <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: '#B23A2E' }}>-{s.qty}</td>
                <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontSize: 11.5, color: '#8A8F97' }}>{s.previousQty} &rarr; {s.newQty}</td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={5} style={{ ...td(), textAlign: 'center', color: '#8A8F97', padding: 26 }}>No sales recorded for this filter yet.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ShippedView({ shippedLog }) {
  const [skuFilter, setSkuFilter] = useState('');
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [sortMode, setSortMode] = useState('date');

  const rows = useMemo(() => {
    const filtered = shippedLog
      .filter(s => !skuFilter || s.sku.toLowerCase().includes(skuFilter.toLowerCase()) || (s.dealer || '').toLowerCase().includes(skuFilter.toLowerCase()) || (s.po || '').toLowerCase().includes(skuFilter.toLowerCase()))
      .filter(s => {
        const d = parseDate(s.date);
        if (!d) return true;
        if (from && d < new Date(from + 'T00:00:00')) return false;
        if (to && d > new Date(to + 'T23:59:59')) return false;
        return true;
      });
    if (sortMode === 'most') return filtered.sort((a, b) => b.qty - a.qty);
    if (sortMode === 'least') return filtered.sort((a, b) => a.qty - b.qty);
    return filtered.sort((a, b) => (parseDate(b.date)?.getTime() ?? 0) - (parseDate(a.date)?.getTime() ?? 0));
  }, [shippedLog, skuFilter, from, to, sortMode]);

  const totalShipped = rows.reduce((s, r) => s + r.qty, 0);

  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Shipped</div>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 8 }}>
        A record of every "Ship" action, kept even after the order itself is fully fulfilled. <strong>{totalShipped}</strong> unit{totalShipped === 1 ? '' : 's'} shipped in the selected range.
      </div>

      <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 8, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: '#8A8F97' }} />
          <input value={skuFilter} onChange={e => setSkuFilter(e.target.value)} placeholder="Filter by model, dealer, or PO…" style={{
            padding: '6px 28px 6px 26px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, width: 240
          }} />
          {skuFilter && <ClearButton onClick={() => setSkuFilter('')} />}
        </div>
        <label style={{ fontSize: 12, color: '#5B6470', display: 'flex', alignItems: 'center', gap: 5 }}>
          From
          <input type="date" value={from} onChange={e => setFrom(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 12.5 }} />
        </label>
        <label style={{ fontSize: 12, color: '#5B6470', display: 'flex', alignItems: 'center', gap: 5 }}>
          To
          <input type="date" value={to} onChange={e => setTo(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 12.5 }} />
        </label>
        <select value={sortMode} onChange={e => setSortMode(e.target.value)} style={{ padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 12.5 }}>
          <option value="date">Sort: newest first</option>
          <option value="most">Sort: most shipped</option>
          <option value="least">Sort: least shipped</option>
        </select>
        {(from || to || skuFilter) && (
          <button onClick={() => { setFrom(''); setTo(''); setSkuFilter(''); }} style={{
            fontSize: 12, color: '#B23A2E', background: 'none', border: 'none', fontWeight: 600
          }}>Clear filters</button>
        )}
      </div>

      <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'hidden' }}>
        <table>
          <thead>
            <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
              <th style={th()}>Date</th>
              <th style={th()}>Model</th>
              <th style={th()}>Dealer</th>
              <th style={th()}>PO</th>
              <th style={th()}>Ships from</th>
              <th style={{ ...th(), textAlign: 'right' }}>Shipped</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((s, i) => (
              <tr key={s.id} className="row-hover" style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                <td style={{ ...td(), fontSize: 12, color: '#5B6470' }}>{s.date}</td>
                <td style={td()}><SkuTag sku={s.sku} /></td>
                <td style={td()}>{s.dealer}</td>
                <td style={{ ...td(), color: '#8A8F97' }}>{s.po || '—'}</td>
                <td style={td()}><Badge color={LOCATIONS[s.shipFrom].color} bg="transparent" border={LOCATIONS[s.shipFrom].color}>{s.shipFrom}</Badge></td>
                <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, color: '#3E7B4F' }}>{s.qty}</td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={6} style={{ ...td(), textAlign: 'center', color: '#8A8F97', padding: 26 }}>No shipments recorded for this filter yet.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ModelsView({ inventory, orders, pendingBySku, onAdd, onRename, onDelete, canEdit }) {
  const [search, setSearch] = useState('');
  const [newSku, setNewSku] = useState('');
  const [editingSku, setEditingSku] = useState(null);
  const [renameValue, setRenameValue] = useState('');
  const [confirmDeleteSku, setConfirmDeleteSku] = useState(null);

  const rows = inventory
    .filter(r => r.sku.toLowerCase().includes(search.toLowerCase()))
    .sort((a, b) => a.sku.localeCompare(b.sku));

  function openOrderCount(sku) {
    return orders.filter(o => o.sku === sku && o.backordered > 0).length;
  }
  function saveRename(oldSku) {
    if (renameValue.trim() && renameValue.trim() !== oldSku) onRename(oldSku, renameValue.trim());
    setEditingSku(null);
  }
  function submitAdd() {
    if (newSku.trim()) { onAdd(newSku); setNewSku(''); }
  }

  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Models</div>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 8 }}>
        Add, rename, or remove bumper models from the catalog. Renaming a model updates every order and log entry that references it, so nothing gets orphaned.
      </div>

      <div style={{ display: 'flex', gap: 8, marginBottom: 10, alignItems: 'center', flexWrap: 'wrap' }}>
        <div style={{ position: 'relative' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 10, color: '#8A8F97' }} />
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search model…" style={{
            padding: '6px 28px 6px 26px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, width: 220
          }} />
          {search && <ClearButton onClick={() => setSearch('')} />}
        </div>
        <div style={{ display: 'flex', gap: 6, marginLeft: 'auto' }}>
          {canEdit && (
            <>
              <input value={newSku} onChange={e => setNewSku(e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') submitAdd(); }}
                placeholder="New model code (e.g. FBMC 03-07)" style={{
                  padding: '6px 8px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, width: 220
                }} />
              <button onClick={submitAdd} style={{
                display: 'flex', alignItems: 'center', gap: 6, background: '#E8592A', color: 'white',
                border: 'none', borderRadius: 7, padding: '7px 12px', fontSize: 12.5, fontWeight: 700
              }}><Plus size={13} /> Add Model</button>
            </>
          )}
        </div>
      </div>

      <div style={{ background: 'white', borderRadius: 10, border: '1px solid #DCD9CE', overflow: 'hidden' }}>
        <table>
          <thead>
            <tr style={{ background: '#1C2126', color: '#F5F3EE' }}>
              <th style={th()}>Model</th>
              <th style={{ ...th(), textAlign: 'right' }}>MX Qty</th>
              <th style={{ ...th(), textAlign: 'right' }}>US Qty</th>
              <th style={{ ...th(), textAlign: 'right' }}>In Production</th>
              <th style={th()}></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r, i) => {
              const isEditing = editingSku === r.sku;
              const isConfirmingDelete = confirmDeleteSku === r.sku;
              const pending = pendingBySku[r.sku] || 0;
              return (
                <tr key={r.sku} className="row-hover" style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                  <td style={td()}>
                    {isEditing ? (
                      <input autoFocus value={renameValue} onChange={e => setRenameValue(e.target.value)}
                        onKeyDown={e => { if (e.key === 'Enter') saveRename(r.sku); if (e.key === 'Escape') setEditingSku(null); }}
                        style={{ padding: '3px 6px', borderRadius: 5, border: '1px solid #33546E', fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, width: 150 }} />
                    ) : (
                      <SkuTag sku={r.sku} />
                    )}
                  </td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace" }}>{r.mx}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace" }}>{r.us}</td>
                  <td style={{ ...td(), textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace", color: pending > 0 ? '#B58A2E' : '#C9C5B8' }}>{pending || '—'}</td>
                  <td style={{ ...td(), textAlign: 'right' }}>
                    {isConfirmingDelete ? (
                      <span style={{ display: 'flex', gap: 6, justifyContent: 'flex-end', alignItems: 'center', fontSize: 11.5 }}>
                        <span style={{ color: '#B23A2E' }}>
                          {openOrderCount(r.sku) > 0 ? `${openOrderCount(r.sku)} open order(s) — delete anyway?` : 'Delete this model?'}
                        </span>
                        <button onClick={() => { onDelete(r.sku); setConfirmDeleteSku(null); }} style={{ background: '#B23A2E', color: 'white', border: 'none', borderRadius: 5, padding: '3px 8px', fontSize: 11, fontWeight: 700 }}>Yes</button>
                        <button onClick={() => setConfirmDeleteSku(null)} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 5, padding: '3px 8px', fontSize: 11 }}>Cancel</button>
                      </span>
                    ) : isEditing ? (
                      <span style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                        <button onClick={() => saveRename(r.sku)} style={{ background: '#3E7B4F', color: 'white', border: 'none', borderRadius: 5, padding: '3px 8px', fontSize: 11, fontWeight: 700 }}>Save</button>
                        <button onClick={() => setEditingSku(null)} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 5, padding: '3px 8px', fontSize: 11 }}>Cancel</button>
                      </span>
                    ) : canEdit ? (
                      <span style={{ display: 'flex', gap: 6, justifyContent: 'flex-end' }}>
                        <button onClick={() => { setEditingSku(r.sku); setRenameValue(r.sku); }} style={{ background: '#EAF0F4', border: '1px solid #C7D6DE', color: '#33546E', borderRadius: 5, padding: '3px 8px', fontSize: 11, fontWeight: 700, display: 'flex', alignItems: 'center', gap: 4 }}><Edit3 size={11} /> Rename</button>
                        <button onClick={() => setConfirmDeleteSku(r.sku)} style={{ background: '#FCEEE8', border: '1px solid #F0C4B8', color: '#B23A2E', borderRadius: 5, padding: '3px 8px', fontSize: 11, fontWeight: 700 }}>Delete</button>
                      </span>
                    ) : null}
                  </td>
                </tr>
              );
            })}
            {rows.length === 0 && (
              <tr><td colSpan={5} style={{ ...td(), textAlign: 'center', color: '#8A8F97', padding: 26 }}>No models match.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ResetDataView({ onReset }) {
  const [code, setCode] = useState('');
  const [unlocked, setUnlocked] = useState(false);
  const [error, setError] = useState(false);
  const [checking, setChecking] = useState(false);
  const [selected, setSelected] = useState({});
  const [confirming, setConfirming] = useState(false);
  const [confirmText, setConfirmText] = useState('');

  const CATEGORIES = [
    { id: 'models', label: 'Models & Inventory' },
    { id: 'dealers', label: 'Dealers' },
    { id: 'orders', label: 'Orders' },
    { id: 'sold', label: 'Sold Units log' },
    { id: 'production', label: 'Production tracking' },
    { id: 'shipped', label: 'Shipped log' },
  ];

  async function tryUnlock() {
    setChecking(true);
    setError(false);
    try {
      const res = await fetch('/api/admin-auth', {
        method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'login', password: code }),
      });
      if (res.ok) setUnlocked(true);
      else setError(true);
    } catch (e) {
      // No backend reachable — likely the Claude preview, which has no real password.
      // Fall back to a simple demo code so this section is still testable here.
      if (code === '0000') setUnlocked(true);
      else setError(true);
    }
    setChecking(false);
  }

  function toggle(id) {
    setSelected(s => ({ ...s, [id]: !s[id] }));
  }

  const selectedIds = CATEGORIES.filter(c => selected[c.id]).map(c => c.id);

  function handleDeleteClick() {
    if (selectedIds.length === 0) return;
    setConfirming(true);
    setConfirmText('');
  }

  function handleConfirm() {
    onReset(selectedIds);
    setSelected({});
    setConfirming(false);
    setConfirmText('');
  }

  if (!unlocked) {
    return (
      <div>
        <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Reset Data</div>
        <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 14 }}>Re-enter your admin password to unlock this section.</div>
        <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 20, maxWidth: 300 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12, color: '#5B6470' }}>
            <Lock size={16} />
            <span style={{ fontSize: 13, fontWeight: 600 }}>Locked</span>
          </div>
          <input type="password" value={code} onChange={e => { setCode(e.target.value); setError(false); }}
            onKeyDown={e => { if (e.key === 'Enter') tryUnlock(); }}
            placeholder="Admin password" style={{
              width: '100%', padding: '8px 10px', borderRadius: 7, border: `1px solid ${error ? '#B23A2E' : '#DCD9CE'}`,
              fontSize: 14, marginBottom: 8
            }} />
          {error && <div style={{ color: '#B23A2E', fontSize: 12, marginBottom: 8 }}>Incorrect password.</div>}
          <button disabled={checking} onClick={tryUnlock} style={{ width: '100%', background: '#1C2126', color: 'white', border: 'none', borderRadius: 7, padding: '9px', fontSize: 13, fontWeight: 700 }}>{checking ? 'Checking…' : 'Unlock'}</button>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Reset Data</div>
      <div style={{ fontSize: 12.5, color: '#B23A2E', marginBottom: 14 }}>
        Check what you want to permanently delete. This can't be undone.
      </div>
      <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: '4px 16px', maxWidth: 380, marginBottom: 14 }}>
        {CATEGORIES.map((c, i) => (
          <label key={c.id} style={{
            display: 'flex', alignItems: 'center', gap: 10, padding: '9px 0', fontSize: 13.5, cursor: 'pointer',
            borderTop: i ? '1px solid #EFEDE4' : 'none'
          }}>
            <input type="checkbox" checked={!!selected[c.id]} onChange={() => toggle(c.id)} style={{ width: 15, height: 15 }} />
            {c.label}
          </label>
        ))}
      </div>

      {!confirming ? (
        <button disabled={selectedIds.length === 0} onClick={handleDeleteClick} style={{
          display: 'flex', alignItems: 'center', gap: 6,
          background: selectedIds.length ? '#B23A2E' : '#DCD9CE', color: 'white', border: 'none',
          borderRadius: 7, padding: '9px 14px', fontSize: 13, fontWeight: 700
        }}><Trash2 size={14} /> Delete Selected</button>
      ) : (
        <div style={{ background: '#FCEEE8', border: '1px solid #F0C4B8', borderRadius: 10, padding: 16, maxWidth: 380 }}>
          <div style={{ fontSize: 13, color: '#B23A2E', marginBottom: 10, fontWeight: 600 }}>
            This will permanently delete: {selectedIds.map(id => CATEGORIES.find(c => c.id === id).label).join(', ')}.
          </div>
          <div style={{ fontSize: 12, color: '#5B6470', marginBottom: 8 }}>Type DELETE to confirm.</div>
          <input value={confirmText} onChange={e => setConfirmText(e.target.value)} style={{
            width: '100%', padding: '8px 10px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, marginBottom: 10
          }} />
          <div style={{ display: 'flex', gap: 8 }}>
            <button disabled={confirmText !== 'DELETE'} onClick={handleConfirm} style={{
              flex: 1, background: confirmText === 'DELETE' ? '#B23A2E' : '#DCD9CE', color: 'white', border: 'none',
              borderRadius: 7, padding: '9px', fontSize: 13, fontWeight: 700
            }}>Confirm Delete</button>
            <button onClick={() => setConfirming(false)} style={{ flex: 1, background: 'white', border: '1px solid #DCD9CE', borderRadius: 7, padding: '9px', fontSize: 13 }}>Cancel</button>
          </div>
        </div>
      )}
    </div>
  );
}

function NewOrdersView({ newOrders, onProcess, onReject, canProcess }) {
  const [poDrafts, setPoDrafts] = useState({});

  const groups = useMemo(() => {
    const map = {};
    newOrders.forEach(o => {
      const key = `${o.dealer}|${o.date}|${o.po}`;
      if (!map[key]) map[key] = { key, dealer: o.dealer, date: o.date, po: o.po, items: [] };
      map[key].items.push(o);
    });
    return Object.values(map).sort((a, b) => (parseDate(b.date)?.getTime() ?? 0) - (parseDate(a.date)?.getTime() ?? 0));
  }, [newOrders]);

  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>New Orders</div>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 12 }}>
        Orders dealers placed through the portal, waiting for you to capture them. Nothing here shows up anywhere else — Orders, Bumper Lookup, Production Planning — until you process it. A PO number is required before you can process one.
      </div>

      {groups.length === 0 ? (
        <div style={{ textAlign: 'center', padding: 50, color: '#5B6470', background: 'white', border: '1px solid #DCD9CE', borderRadius: 10 }}>
          <PackageCheck size={28} color="#3E7B4F" style={{ marginBottom: 8 }} />
          <div style={{ fontSize: 13 }}>No new order requests from dealers right now.</div>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {groups.map(g => {
            const poValue = poDrafts[g.key] ?? (g.po || '');
            const poFilled = poValue.trim().length > 0;
            return (
              <div key={g.key} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 14 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, flexWrap: 'wrap', gap: 8 }}>
                  <div>
                    <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 14, textTransform: 'uppercase' }}>{g.dealer}</div>
                    <div style={{ fontSize: 11.5, color: '#8A8F97' }}>
                      {g.date || '—'} · {g.items.length} model{g.items.length === 1 ? '' : 's'}
                    </div>
                  </div>
                  {canProcess ? (
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                      <input
                        value={poValue}
                        onChange={e => setPoDrafts(d => ({ ...d, [g.key]: e.target.value }))}
                        placeholder="PO # (required)"
                        style={{
                          padding: '6px 9px', borderRadius: 6, border: `1px solid ${poFilled ? '#DCD9CE' : '#E8B4A8'}`,
                          fontSize: 12.5, width: 140
                        }}
                      />
                      <button onClick={() => onReject(g.items.map(o => o.id))} style={{
                        fontSize: 12, fontWeight: 700, color: '#B23A2E', background: '#FCEEE8', border: '1px solid #F0C4B8',
                        borderRadius: 6, padding: '6px 10px'
                      }}>Reject</button>
                      <button
                        disabled={!poFilled}
                        onClick={() => onProcess(g.items.map(o => o.id), poValue.trim())}
                        title={poFilled ? '' : 'Enter a PO number first'}
                        style={{
                          display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, fontWeight: 700, color: 'white',
                          background: poFilled ? '#E8592A' : '#DCD9CE', border: 'none', borderRadius: 6, padding: '6px 12px'
                        }}><Check size={12} /> Process order</button>
                    </div>
                  ) : (
                    <span style={{ fontSize: 11.5, color: '#8A8F97', fontStyle: 'italic' }}>View only</span>
                  )}
                  </div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                  {g.items.map(o => (
                    <div key={o.id} style={{ display: 'flex', alignItems: 'center', gap: 6, background: '#F7F5EF', border: '1px solid #EFEDE4', borderRadius: 7, padding: '4px 8px' }}>
                      <SkuTag sku={o.sku} />
                      <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 700, fontSize: 12.5, color: '#5B6470' }}>× {o.qty}</span>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function BackupView({ state, onRestoreFile }) {
  const [autoBackups, setAutoBackups] = useState(null);
  const [loadingAuto, setLoadingAuto] = useState(true);
  const [restoring, setRestoring] = useState(null);
  const [fileError, setFileError] = useState('');
  const [pendingRestore, setPendingRestore] = useState(null);
  const fileInputRef = React.useRef(null);

  function adminFetch(body) {
    const adminToken = localStorage.getItem('gr-admin-token') || '';
    return fetch('/api/backup-manage', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...body, token: adminToken }),
    });
  }

  useEffect(() => {
    (async () => {
      try {
        const res = await adminFetch({ action: 'list' });
        if (res.ok) {
          const data = await res.json();
          setAutoBackups(data.backups || []);
        } else {
          setAutoBackups([]);
        }
      } catch (e) {
        setAutoBackups([]); // likely the Claude preview, no server to reach
      }
      setLoadingAuto(false);
    })();
  }, []);

  function downloadBackup() {
    const payload = { exportedAt: new Date().toISOString(), data: state };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const stamp = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `gr-bumpers-backup-${stamp}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function handleFileChange(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFileError('');
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed.data) throw new Error('missing data');
        setPendingRestore(parsed);
      } catch (err) {
        setFileError('Could not read that file — make sure it\u2019s a backup exported from this app.');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  }

  async function restoreAutoBackup(date) {
    setRestoring(date);
    try {
      const res = await adminFetch({ action: 'restore', date });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Restore failed.');
      window.location.reload();
    } catch (err) {
      setFileError(err.message || 'Could not restore that backup.');
    }
    setRestoring(null);
  }

  const counts = state ? {
    Models: state.inventory?.length || 0,
    Dealers: state.dealers?.length || 0,
    Orders: state.orders?.length || 0,
    'Sales log entries': state.salesLog?.length || 0,
  } : {};

  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Backup</div>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 14 }}>
        A copy of your data lives here — download one whenever you want, and the site also takes an automatic snapshot once a day so there's always a recent point to roll back to.
      </div>

      <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 16, marginBottom: 14, maxWidth: 460 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
          <DownloadCloud size={15} color="#33546E" />
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 13, textTransform: 'uppercase' }}>Download a backup now</div>
        </div>
        <div style={{ fontSize: 11.5, color: '#8A8F97', marginBottom: 10 }}>
          {Object.entries(counts).map(([k, v]) => `${v} ${k}`).join(' · ')}
        </div>
        <button onClick={downloadBackup} style={{
          display: 'flex', alignItems: 'center', gap: 6, background: '#1C2126', color: 'white', border: 'none',
          borderRadius: 7, padding: '8px 14px', fontSize: 12.5, fontWeight: 700
        }}><DownloadCloud size={13} /> Download backup file</button>
      </div>

      <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 16, marginBottom: 14, maxWidth: 460 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
          <UploadCloud size={15} color="#B58A2E" />
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 13, textTransform: 'uppercase' }}>Restore from a downloaded file</div>
        </div>
        <div style={{ fontSize: 11.5, color: '#8A8F97', marginBottom: 10 }}>
          This replaces your current data with whatever is in the file. Use this if something went wrong and you need to roll back.
        </div>
        <input ref={fileInputRef} type="file" accept=".json" onChange={handleFileChange} style={{ display: 'none' }} />
        <button onClick={() => fileInputRef.current?.click()} style={{
          display: 'flex', alignItems: 'center', gap: 6, background: 'white', border: '1px solid #DCD9CE', color: '#5B6470',
          borderRadius: 7, padding: '8px 14px', fontSize: 12.5, fontWeight: 700
        }}><UploadCloud size={13} /> Choose backup file…</button>
        {fileError && <div style={{ color: '#B23A2E', fontSize: 12, marginTop: 8 }}>{fileError}</div>}
        {pendingRestore && (
          <div style={{ marginTop: 10, background: '#FCEEE8', border: '1px solid #F0C4B8', borderRadius: 8, padding: 12 }}>
            <div style={{ fontSize: 12.5, color: '#B23A2E', marginBottom: 8 }}>
              Restore this backup from {pendingRestore.exportedAt ? new Date(pendingRestore.exportedAt).toLocaleString() : 'unknown date'}? This replaces your current data.
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={() => { onRestoreFile(pendingRestore); setPendingRestore(null); }} style={{
                flex: 1, background: '#B23A2E', color: 'white', border: 'none', borderRadius: 7, padding: '8px', fontSize: 12.5, fontWeight: 700
              }}>Restore</button>
              <button onClick={() => setPendingRestore(null)} style={{
                flex: 1, background: 'white', border: '1px solid #DCD9CE', borderRadius: 7, padding: '8px', fontSize: 12.5
              }}>Cancel</button>
            </div>
          </div>
        )}
      </div>

      <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 16, maxWidth: 460 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
          <RotateCcw size={15} color="#5B6470" />
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 13, textTransform: 'uppercase' }}>Automatic daily snapshots</div>
        </div>
        {loadingAuto ? (
          <div style={{ fontSize: 12, color: '#8A8F97' }}>Checking…</div>
        ) : !autoBackups || autoBackups.length === 0 ? (
          <div style={{ fontSize: 12, color: '#8A8F97' }}>
            None yet — these only exist once deployed live and the daily backup has run at least once.
          </div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            {autoBackups.map(b => (
              <div key={b.date} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', fontSize: 12.5, padding: '6px 0', borderTop: '1px solid #EFEDE4' }}>
                <span style={{ color: '#5B6470' }}>{b.date}</span>
                <button disabled={restoring === b.date} onClick={() => restoreAutoBackup(b.date)} style={{
                  fontSize: 11.5, fontWeight: 700, color: '#33546E', background: '#EAF0F4', border: '1px solid #C7D6DE',
                  borderRadius: 5, padding: '4px 9px'
                }}>{restoring === b.date ? 'Restoring…' : 'Restore this'}</button>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function StaffView() {
  const [staff, setStaff] = useState(null);
  const [loading, setLoading] = useState(true);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [perms, setPerms] = useState({});
  const [error, setError] = useState('');
  const [saving, setSaving] = useState(false);

  const PERMISSION_FIELDS = [
    { id: 'processNewOrders', label: 'Process new orders (dealer requests)' },
    { id: 'createOrders', label: 'Create new orders' },
    { id: 'editOrders', label: 'Edit orders' },
    { id: 'editModelsAndDealers', label: 'Edit models and dealers' },
    { id: 'receiveOrders', label: 'Receive orders (Ship button)' },
    { id: 'manageDealerLogins', label: 'Manage dealer portal logins' },
  ];

  function adminFetch(body) {
    const token = localStorage.getItem('gr-admin-token') || '';
    return fetch('/api/admin-auth', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ...body, token }),
    });
  }

  async function loadStaff() {
    setLoading(true);
    try {
      const res = await adminFetch({ action: 'list-staff' });
      if (res.ok) {
        const data = await res.json();
        setStaff(data.staff || []);
      } else {
        setStaff([]);
      }
    } catch (e) {
      setStaff([]); // likely the Claude preview, no server to reach
    }
    setLoading(false);
  }

  useEffect(() => { loadStaff(); }, []);

  async function addStaff(e) {
    e.preventDefault();
    setError('');
    if (!username.trim() || !password) { setError('Username and password are required.'); return; }
    setSaving(true);
    try {
      const res = await adminFetch({ action: 'create-staff', username: username.trim(), password, permissions: perms });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Could not create staff account.');
      setUsername(''); setPassword(''); setPerms({});
      await loadStaff();
    } catch (err) {
      setError(err.message || 'Could not reach the server — this only works on the live deployed site.');
    }
    setSaving(false);
  }

  async function togglePermission(member, field) {
    const next = { ...(member.permissions || {}), [field]: !member.permissions?.[field] };
    setStaff(s => s.map(m => m.username === member.username ? { ...m, permissions: next } : m));
    try {
      await adminFetch({ action: 'update-staff-permissions', username: member.username, permissions: next });
    } catch (e) { /* best effort */ }
  }

  async function removeStaff(username) {
    setStaff(s => s.filter(m => m.username !== username));
    try {
      await adminFetch({ action: 'delete-staff', username });
    } catch (e) { /* best effort */ }
  }

  return (
    <div>
      <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 15, textTransform: 'uppercase', marginBottom: 4 }}>Staff</div>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 14 }}>
        By default a staff account can only browse — nothing checked means view-only. Turn on whichever specific abilities they need below. Only you (the owner) can see this section.
      </div>

      <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 16, marginBottom: 14, maxWidth: 460 }}>
        <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 13, textTransform: 'uppercase', marginBottom: 10 }}>Add a staff account</div>
        <form onSubmit={addStaff}>
          <label style={labelStyle()}>Username</label>
          <input value={username} onChange={e => setUsername(e.target.value)} style={fieldStyle()} />
          <label style={labelStyle()}>Password</label>
          <input value={password} onChange={e => setPassword(e.target.value)} style={fieldStyle()} />
          <label style={labelStyle()}>Abilities</label>
          <div style={{ marginBottom: 14 }}>
            {PERMISSION_FIELDS.map(f => (
              <label key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, marginBottom: 6, cursor: 'pointer' }}>
                <input type="checkbox" checked={!!perms[f.id]} onChange={e => setPerms(p => ({ ...p, [f.id]: e.target.checked }))} style={{ width: 15, height: 15 }} />
                {f.label}
              </label>
            ))}
          </div>
          {error && <div style={{ color: '#B23A2E', fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
          <button disabled={saving} type="submit" style={{
            width: '100%', background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
            padding: '9px', fontSize: 13.5, fontWeight: 700
          }}>{saving ? 'Adding…' : 'Add staff account'}</button>
        </form>
      </div>

      <div style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 10, padding: 16, maxWidth: 520 }}>
        <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 13, textTransform: 'uppercase', marginBottom: 10 }}>Existing staff</div>
        {loading ? (
          <div style={{ fontSize: 12, color: '#8A8F97' }}>Loading…</div>
        ) : !staff || staff.length === 0 ? (
          <div style={{ fontSize: 12, color: '#8A8F97' }}>No staff accounts yet.</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            {staff.map(m => (
              <div key={m.username} style={{ borderTop: '1px solid #EFEDE4', paddingTop: 10 }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 6 }}>
                  <span style={{ fontSize: 13, fontWeight: 600 }}>{m.username}</span>
                  <button onClick={() => removeStaff(m.username)} style={{
                    background: '#FCEEE8', border: '1px solid #F0C4B8', color: '#B23A2E', borderRadius: 5, padding: '3px 7px', display: 'flex', alignItems: 'center'
                  }}><X size={11} /></button>
                </div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px 14px' }}>
                  {PERMISSION_FIELDS.map(f => (
                    <label key={f.id} style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 11.5, color: '#5B6470', cursor: 'pointer' }}>
                      <input type="checkbox" checked={!!m.permissions?.[f.id]} onChange={() => togglePermission(m, f.id)} style={{ width: 13, height: 13 }} />
                      {f.label}
                    </label>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function ModalShell({ title, onClose, children, width = 420 }) {
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(28,33,38,0.5)', display: 'flex',
      alignItems: 'center', justifyContent: 'center', zIndex: 900, padding: 16
    }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{
        background: 'white', borderRadius: 12, width, maxWidth: '100%', padding: 16,
        boxShadow: '0 20px 60px rgba(0,0,0,0.3)', fontFamily: "'Inter', sans-serif"
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 600, fontSize: 16, textTransform: 'uppercase' }}>{title}</div>
          <button onClick={onClose} style={{ border: 'none', background: 'none', color: '#8A8F97' }}><X size={18} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

function fieldStyle() {
  return { width: '100%', padding: '7px 9px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13, marginBottom: 8 };
}
function labelStyle() {
  return { fontSize: 11.5, fontWeight: 700, color: '#5B6470', textTransform: 'uppercase', letterSpacing: '0.03em', display: 'block', marginBottom: 5 };
}

function ShipModal({ order, onClose, onShip }) {
  const [qty, setQty] = useState(order.backordered);
  return (
    <ModalShell title="Ship order" onClose={onClose}>
      <div style={{ fontSize: 13, marginBottom: 8, color: '#5B6470' }}>
        <SkuTag sku={order.sku} /> to <strong style={{ color: '#1C2126' }}>{order.dealer}</strong> from the <strong>{LOCATIONS[order.shipFrom].label}</strong>
      </div>
      <label style={labelStyle()}>Quantity to ship (owed: {order.backordered})</label>
      <input type="number" value={qty} min={1} max={order.backordered} onChange={e => setQty(parseInt(e.target.value, 10) || 0)} style={fieldStyle()} />
      <button onClick={() => onShip(order, qty)} style={{
        width: '100%', background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
        padding: '9px', fontSize: 13.5, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6
      }}><Check size={15} /> Confirm shipment</button>
    </ModalShell>
  );
}

function OrderModal({ dealers, inventory, onClose, onAdd }) {
  const [dealer, setDealer] = useState('');
  const [po, setPo] = useState('');
  const [date, setDate] = useState('');
  const [lines, setLines] = useState([{ sku: inventory[0]?.sku || '', qty: 1 }]);

  function updateLine(i, updates) {
    setLines(ls => ls.map((l, idx) => idx === i ? { ...l, ...updates } : l));
  }
  function addLine() {
    setLines(ls => [...ls, { sku: inventory[0]?.sku || '', qty: 1 }]);
  }
  function removeLine(i) {
    setLines(ls => ls.filter((_, idx) => idx !== i));
  }

  function submit() {
    const validLines = lines.filter(l => l.sku && l.qty > 0);
    if (!dealer || !po.trim() || validLines.length === 0) return;
    const items = validLines.map(l => ({
      sku: l.sku, dealer, qty: l.qty, po: po.trim(), date: date || new Date().toLocaleDateString('en-US'), due: '', num: '',
    }));
    onAdd(items);
  }

  const canSubmit = !!dealer && !!po.trim() && lines.some(l => l.sku && l.qty > 0);

  return (
    <ModalShell title="New order" onClose={onClose} width={460}>
      <label style={labelStyle()}>Dealer</label>
      <input list="dealer-list" value={dealer} onChange={e => setDealer(e.target.value)} placeholder="Start typing…" style={fieldStyle()} />
      <datalist id="dealer-list">
        {dealers.map(d => <option key={d.name} value={d.name} />)}
      </datalist>
      <label style={labelStyle()}>PO # (required, applies to the whole order)</label>
      <input value={po} onChange={e => setPo(e.target.value)} style={{ ...fieldStyle(), border: `1px solid ${po.trim() ? '#DCD9CE' : '#E8B4A8'}` }} />

      <label style={labelStyle()}>Models</label>
      {lines.map((l, i) => (
        <div key={i} style={{ display: 'flex', gap: 8, marginBottom: 8, alignItems: 'center' }}>
          <select value={l.sku} onChange={e => updateLine(i, { sku: e.target.value })} style={{ ...fieldStyle(), marginBottom: 0, flex: 2 }}>
            {inventory.map(inv => <option key={inv.sku} value={inv.sku}>{inv.sku}</option>)}
          </select>
          <input type="number" value={l.qty} min={1} onChange={e => updateLine(i, { qty: parseInt(e.target.value, 10) || 1 })} style={{ ...fieldStyle(), marginBottom: 0, width: 70, flex: 'none' }} />
          {lines.length > 1 && (
            <button onClick={() => removeLine(i)} style={{
              background: '#FCEEE8', border: '1px solid #F0C4B8', color: '#B23A2E', borderRadius: 6,
              padding: '7px 9px', display: 'flex', alignItems: 'center', flex: 'none'
            }}><X size={13} /></button>
          )}
        </div>
      ))}
      <button onClick={addLine} style={{
        display: 'flex', alignItems: 'center', gap: 6, background: 'white', border: '1px dashed #C9C5B8', color: '#5B6470',
        borderRadius: 7, padding: '7px 10px', fontSize: 12.5, fontWeight: 600, marginBottom: 16, width: '100%', justifyContent: 'center'
      }}><Plus size={13} /> Add another model</button>

      <button disabled={!canSubmit} onClick={submit} title={canSubmit ? '' : 'Dealer, PO, and at least one model are required'} style={{
        width: '100%', background: canSubmit ? '#E8592A' : '#DCD9CE', color: 'white', border: 'none', borderRadius: 8,
        padding: '9px', fontSize: 13.5, fontWeight: 700
      }}>Add order{lines.filter(l => l.sku && l.qty > 0).length > 1 ? ` (${lines.filter(l => l.sku && l.qty > 0).length} models)` : ''}</button>
    </ModalShell>
  );
}

function DealerModal({ onClose, onAdd }) {
  const [name, setName] = useState('');
  const [origin, setOrigin] = useState('MX');
  return (
    <ModalShell title="New dealer" onClose={onClose}>
      <label style={labelStyle()}>Dealer name</label>
      <input value={name} onChange={e => setName(e.target.value)} style={fieldStyle()} />
      <label style={labelStyle()}>Ships from</label>
      <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
        {['MX', 'US'].map(o => (
          <button key={o} onClick={() => setOrigin(o)} style={{
            flex: 1, padding: '9px', borderRadius: 7, fontSize: 12.5, fontWeight: 700,
            border: origin === o ? `2px solid ${LOCATIONS[o].color}` : '1px solid #DCD9CE',
            background: origin === o ? '#F7F5EF' : 'white'
          }}>{LOCATIONS[o].label}</button>
        ))}
      </div>
      <button onClick={() => name && onAdd({ name, origin })} style={{
        width: '100%', background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
        padding: '9px', fontSize: 13.5, fontWeight: 700
      }}>Add dealer</button>
    </ModalShell>
  );
}

function ImportModal({ location, inventory, pendingBySku, onClose, onApply }) {
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState(null);
  const [fileName, setFileName] = useState(null);
  const locKey = LOCATIONS[location].key;

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setError(null);
    setPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
      const { matched, unmatched } = parseInventoryRows(rows, location === 'US');
      if (matched.length === 0) {
        setError('No stock rows recognized in that file. Make sure it\u2019s the right export.');
        return;
      }
      setPreview({ matched, unmatched });
    } catch (err) {
      setError('Could not read that file — make sure it\u2019s a .xlsx export.');
    }
  }

  function apply() {
    if (!preview) return;
    const bySku = {};
    preview.matched.forEach(m => { bySku[m.sku] = (bySku[m.sku] || 0) + m.qty; });
    let next = inventory.map(row => ({ ...row, [locKey]: bySku[row.sku] !== undefined ? bySku[row.sku] : 0 }));
    Object.keys(bySku).forEach(sku => {
      if (!next.some(r => r.sku === sku)) {
        next.push({ sku, mx: 0, us: 0, [locKey]: bySku[sku] });
      }
    });

    // A drop in on-hand qty for a SKU we already had on record = units shipped/sold since the last import
    const todayStr = formatToday();
    const sales = [];
    const productionReceipts = [];
    inventory.forEach(row => {
      const prevQty = row[locKey] || 0;
      const newRow = next.find(r => r.sku === row.sku);
      const newQty = newRow ? (newRow[locKey] || 0) : 0;
      if (prevQty > newQty) {
        sales.push({ sku: row.sku, location, qty: prevQty - newQty, previousQty: prevQty, newQty, date: todayStr });
      } else if (location === 'MX' && newQty > prevQty) {
        productionReceipts.push({ sku: row.sku, qty: newQty - prevQty, previousQty: prevQty, newQty, date: todayStr });
      }
    });

    onApply(next, sales, productionReceipts);
  }

  const zeroedSkus = preview ? inventory.filter(r => !preview.matched.some(m => m.sku === r.sku) && (r[locKey] || 0) > 0) : [];

  const soldPreview = useMemo(() => {
    if (!preview) return 0;
    const bySku = {};
    preview.matched.forEach(m => { bySku[m.sku] = (bySku[m.sku] || 0) + m.qty; });
    let total = 0;
    inventory.forEach(row => {
      const prevQty = row[locKey] || 0;
      const newQty = bySku[row.sku] !== undefined ? bySku[row.sku] : 0;
      if (prevQty > newQty) total += prevQty - newQty;
    });
    return total;
  }, [preview, inventory, locKey]);

  const productionPreview = useMemo(() => {
    if (!preview || location !== 'MX') return 0;
    const bySku = {};
    preview.matched.forEach(m => { bySku[m.sku] = (bySku[m.sku] || 0) + m.qty; });
    let total = 0;
    inventory.forEach(row => {
      const prevQty = row[locKey] || 0;
      const newQty = bySku[row.sku] !== undefined ? bySku[row.sku] : 0;
      if (newQty > prevQty) total += Math.min(newQty - prevQty, pendingBySku[row.sku] || 0);
    });
    return total;
  }, [preview, inventory, locKey, location, pendingBySku]);

  return (
    <ModalShell title={`Import ${LOCATIONS[location].label} stock`} onClose={onClose} width={500}>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 8, lineHeight: 1.5 }}>
        Upload the {location === 'MX' ? '"Existencia y valor del inventario"' : 'GR/US stock'} export (.xlsx). This replaces the current on-hand counts for {LOCATIONS[location].label} with what\u2019s in the file — any model not listed is set to 0, same as "solo existencias" only shows items in stock.
      </div>
      <label style={{
        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, border: '2px dashed #DCD9CE',
        borderRadius: 8, padding: '14px', marginBottom: 8, cursor: 'pointer', fontSize: 13, color: '#5B6470', background: '#FAFAF7'
      }}>
        <Upload size={16} />
        {fileName || 'Click to choose .xlsx file'}
        <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: 'none' }} />
      </label>

      {error && <div style={{ color: '#B23A2E', fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

      {preview && (
        <div style={{ marginBottom: 8 }}>
          <div style={{ fontSize: 12.5, marginBottom: 8 }}>
            <strong>{preview.matched.length}</strong> models will be set from the file
            {zeroedSkus.length > 0 && <> · <strong>{zeroedSkus.length}</strong> current models not in the file will be set to 0</>}
          </div>
          {soldPreview > 0 && (
            <div style={{ fontSize: 12.5, color: '#B23A2E', background: '#FCEEE8', border: '1px solid #F0C4B8', borderRadius: 7, padding: '7px 10px', marginBottom: 8 }}>
              <strong>{soldPreview}</strong> units will be logged as sold (on-hand quantity dropped since the last import)
            </div>
          )}
          {productionPreview > 0 && (
            <div style={{ fontSize: 12.5, color: '#3E7B4F', background: '#EAF4EC', border: '1px solid #C7E0CC', borderRadius: 7, padding: '7px 10px', marginBottom: 8 }}>
              <strong>{productionPreview}</strong> units will be credited against pending production (on-hand quantity increased)
            </div>
          )}

          {preview.unmatched.length > 0 && (
            <div style={{ background: '#FCEEE8', border: '1px solid #F0C4B8', borderRadius: 7, padding: 10, marginBottom: 10 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: '#B23A2E', marginBottom: 4 }}>
                {preview.unmatched.length} item{preview.unmatched.length === 1 ? '' : 's'} couldn't be matched to a model code:
              </div>
              {preview.unmatched.map((u, i) => (
                <div key={i} style={{ fontSize: 11.5, color: '#5B6470' }}>{u.label} — qty {u.qty}</div>
              ))}
            </div>
          )}

          <div style={{ maxHeight: 220, overflowY: 'auto', border: '1px solid #EFEDE4', borderRadius: 7 }}>
            <table>
              <tbody>
                {preview.matched.map((m, i) => (
                  <tr key={i} style={{ borderTop: i ? '1px solid #EFEDE4' : 'none' }}>
                    <td style={{ ...td(), fontSize: 12 }}><SkuTag sku={m.sku} /></td>
                    <td style={{ ...td(), fontSize: 12, textAlign: 'right', fontFamily: "'IBM Plex Mono', monospace" }}>{m.qty}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      <button disabled={!preview} onClick={apply} style={{
        width: '100%', background: preview ? '#E8592A' : '#DCD9CE', color: 'white', border: 'none', borderRadius: 8,
        padding: '9px', fontSize: 13.5, fontWeight: 700
      }}>Apply import</button>
    </ModalShell>
  );
}

function OrdersImportModal({ onClose, onApply }) {
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState(null);
  const [fileName, setFileName] = useState(null);

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    setError(null);
    setPreview(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
      const parsed = parseOrdersRows(rows);
      if (parsed.length === 0) {
        setError("No order rows recognized in that file. Make sure it's the open sales orders export.");
        return;
      }
      const skus = new Set(parsed.map(o => o.sku));
      const dates = parsed.map(o => parseDate(o.date)).filter(Boolean).sort((a, b) => a - b);
      const totalBackordered = parsed.reduce((s, o) => s + o.backordered, 0);
      setPreview({
        rows: parsed, skuCount: skus.size,
        dateRange: dates.length ? [dates[0], dates[dates.length - 1]] : null,
        totalBackordered,
      });
    } catch (err) {
      setError("Could not read that file — make sure it's a .xlsx export.");
    }
  }

  function apply() {
    if (!preview) return;
    const withIds = preview.rows.map((o, i) => ({ id: 'O' + (i + 1), ...o }));
    onApply(withIds);
  }

  function fmt(d) {
    if (!d) return '';
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
  }

  return (
    <ModalShell title="Import orders" onClose={onClose} width={480}>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 8, lineHeight: 1.5 }}>
        Upload the QuickBooks open sales orders export (.xlsx). This replaces the entire order list — every dealer, PO, date, and quantity owed comes from this file.
      </div>
      <label style={{
        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, border: '2px dashed #DCD9CE',
        borderRadius: 8, padding: '14px', marginBottom: 8, cursor: 'pointer', fontSize: 13, color: '#5B6470', background: '#FAFAF7'
      }}>
        <Upload size={16} />
        {fileName || 'Click to choose .xlsx file'}
        <input type="file" accept=".xlsx" onChange={handleFile} style={{ display: 'none' }} />
      </label>

      {error && <div style={{ color: '#B23A2E', fontSize: 12.5, marginBottom: 12 }}>{error}</div>}

      {preview && (
        <div style={{ marginBottom: 8, fontSize: 12.5, color: '#5B6470' }}>
          <div style={{ background: '#FAFAF7', border: '1px solid #EFEDE4', borderRadius: 7, padding: 12, marginBottom: 10 }}>
            <div><strong>{preview.rows.length}</strong> order lines across <strong>{preview.skuCount}</strong> models</div>
            <div><strong>{preview.totalBackordered}</strong> total units currently owed</div>
            {preview.dateRange && <div>Dates from {fmt(preview.dateRange[0])} to {fmt(preview.dateRange[1])}</div>}
          </div>
          <div style={{ color: '#B23A2E' }}>This will replace all {preview.rows.length} orders currently in the system.</div>
        </div>
      )}

      <button disabled={!preview} onClick={apply} style={{
        width: '100%', background: preview ? '#E8592A' : '#DCD9CE', color: 'white', border: 'none', borderRadius: 8,
        padding: '9px', fontSize: 13.5, fontWeight: 700
      }}>Replace orders with this file</button>
    </ModalShell>
  );
}

function DealerLoginModal({ dealerName, onClose }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [confirmPassword, setConfirmPassword] = useState('');
  const [existing, setExisting] = useState(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch('/api/dealer-auth', {
          method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'list', token: localStorage.getItem('gr-admin-token') || '' }),
        });
        const data = await res.json();
        const found = (data.accounts || []).find(a => a.dealerName === dealerName);
        if (found) { setExisting(found.username); setUsername(found.username); }
      } catch (e) {
        // Likely running outside the deployed site (e.g. the Claude preview) — fine, just no data to show.
      }
      setLoading(false);
    })();
  }, [dealerName]);

  function generatePassword() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
    let pw = '';
    for (let i = 0; i < 12; i++) pw += chars[Math.floor(Math.random() * chars.length)];
    setPassword(pw);
    setShowPassword(true);
  }

  async function save() {
    if (!username.trim() || !password) { setError('Username and password are required.'); return; }
    if (existing && !confirmPassword) { setError('Enter your own password to confirm this change.'); return; }
    setSaving(true); setError(''); setSuccess('');
    try {
      const res = await fetch('/api/dealer-auth', {
        method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'set-credentials', dealerName, username: username.trim(), password,
          confirmPassword, token: localStorage.getItem('gr-admin-token') || '',
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Could not save.');
      setSuccess(`Login saved. Password: ${password} — write this down, it won't be shown again.`);
      setExisting(username.trim());
      setPassword('');
      setConfirmPassword('');
      setShowPassword(false);
    } catch (e) {
      setError(e.message || 'Could not reach the server — this only works on the live deployed site.');
    }
    setSaving(false);
  }

  async function remove() {
    setSaving(true); setError(''); setSuccess('');
    try {
      const res = await fetch('/api/dealer-auth', {
        method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'remove-credentials', dealerName, token: localStorage.getItem('gr-admin-token') || '' }),
      });
      if (!res.ok) { const data = await res.json(); throw new Error(data.error || 'Could not remove.'); }
      setExisting(null); setUsername(''); setPassword('');
      setSuccess('Login removed.');
    } catch (e) {
      setError(e.message || 'Could not reach the server.');
    }
    setSaving(false);
  }

  return (
    <ModalShell title="Dealer portal login" onClose={onClose} width={380}>
      <div style={{ fontSize: 12.5, color: '#5B6470', marginBottom: 14 }}>
        Set a username and password so <strong>{dealerName}</strong> can log into the dealer portal to place orders and see their own order history — nothing else.
      </div>
      {loading ? (
        <div style={{ fontSize: 12.5, color: '#8A8F97' }}>Checking current status…</div>
      ) : (
        <>
          {existing && <div style={{ fontSize: 12, color: '#3E7B4F', marginBottom: 10 }}>Currently has a login: <strong>{existing}</strong></div>}
          <label style={labelStyle()}>Username</label>
          <input value={username} onChange={e => setUsername(e.target.value)} style={fieldStyle()} />
          <label style={labelStyle()}>{existing ? 'New password' : 'Password'}</label>
          <div style={{ display: 'flex', gap: 6, marginBottom: 12 }}>
            <input type={showPassword ? 'text' : 'password'} value={password} onChange={e => setPassword(e.target.value)}
              style={{ ...fieldStyle(), marginBottom: 0, flex: 1 }} placeholder={existing ? 'Enter a new password to reset it' : ''} />
            <button type="button" onClick={() => setShowPassword(s => !s)} style={{
              background: 'white', border: '1px solid #DCD9CE', borderRadius: 7, padding: '0 10px', fontSize: 11.5, color: '#5B6470', flexShrink: 0
            }}>{showPassword ? 'Hide' : 'Show'}</button>
          </div>
          <button type="button" onClick={generatePassword} style={{
            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, width: '100%',
            background: '#FAFAF7', border: '1px dashed #C9C5B8', color: '#5B6470', borderRadius: 7,
            padding: '7px 10px', fontSize: 12, fontWeight: 600, marginBottom: 14
          }}><KeyRound size={12} /> Generate a random password</button>

          {existing && (
            <>
              <label style={labelStyle()}>Your admin password (confirm this change)</label>
              <input type="password" value={confirmPassword} onChange={e => setConfirmPassword(e.target.value)} style={fieldStyle()} />
            </>
          )}

          {error && <div style={{ color: '#B23A2E', fontSize: 12.5, marginBottom: 10 }}>{error}</div>}
          {success && <div style={{ color: '#3E7B4F', fontSize: 12.5, marginBottom: 10 }}>{success}</div>}
          <button disabled={saving} onClick={save} style={{
            width: '100%', background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
            padding: '9px', fontSize: 13.5, fontWeight: 700, marginBottom: existing ? 8 : 0
          }}>{saving ? 'Saving…' : existing ? 'Update login' : 'Create login'}</button>
          {existing && (
            <button disabled={saving} onClick={remove} style={{
              width: '100%', background: 'white', border: '1px solid #F0C4B8', color: '#B23A2E',
              borderRadius: 8, padding: '9px', fontSize: 13, fontWeight: 700
            }}>Remove login</button>
          )}
        </>
      )}
    </ModalShell>
  );
}

function EditOrderModal({ order, dealers, inventory, onClose, onSave, onDelete }) {
  const [sku, setSku] = useState(order.sku);
  const [dealer, setDealer] = useState(order.dealer);
  const [po, setPo] = useState(order.po || '');
  const [date, setDate] = useState(order.date || '');
  const [qty, setQty] = useState(order.qty);
  const [invoiced, setInvoiced] = useState(order.invoiced);
  const [confirmingDelete, setConfirmingDelete] = useState(false);

  function save() {
    const cleanQty = Math.max(0, parseInt(qty, 10) || 0);
    const cleanInvoiced = Math.max(0, Math.min(cleanQty, parseInt(invoiced, 10) || 0));
    onSave(order.id, {
      sku, dealer, po, date, qty: cleanQty, invoiced: cleanInvoiced,
      backordered: cleanQty - cleanInvoiced,
    });
    onClose();
  }

  return (
    <ModalShell title="Edit order" onClose={onClose} width={420}>
      <label style={labelStyle()}>Model</label>
      <select value={sku} onChange={e => setSku(e.target.value)} style={fieldStyle()}>
        {inventory.map(i => <option key={i.sku} value={i.sku}>{i.sku}</option>)}
        {!inventory.some(i => i.sku === sku) && <option value={sku}>{sku} (not in catalog)</option>}
      </select>

      <label style={labelStyle()}>Dealer</label>
      <input list="edit-order-dealer-list" value={dealer} onChange={e => setDealer(e.target.value)} style={fieldStyle()} />
      <datalist id="edit-order-dealer-list">
        {dealers.map(d => <option key={d.name} value={d.name} />)}
      </datalist>

      <div style={{ display: 'flex', gap: 10 }}>
        <div style={{ flex: 1 }}>
          <label style={labelStyle()}>PO #</label>
          <input value={po} onChange={e => setPo(e.target.value)} style={fieldStyle()} />
        </div>
        <div style={{ flex: 1 }}>
          <label style={labelStyle()}>Date</label>
          <input value={date} onChange={e => setDate(e.target.value)} placeholder="M/D/YYYY" style={fieldStyle()} />
        </div>
      </div>

      <div style={{ display: 'flex', gap: 10 }}>
        <div style={{ flex: 1 }}>
          <label style={labelStyle()}>Total ordered</label>
          <input type="number" min="0" value={qty} onChange={e => setQty(e.target.value)} style={fieldStyle()} />
        </div>
        <div style={{ flex: 1 }}>
          <label style={labelStyle()}>Already shipped</label>
          <input type="number" min="0" value={invoiced} onChange={e => setInvoiced(e.target.value)} style={fieldStyle()} />
        </div>
      </div>
      <div style={{ fontSize: 11.5, color: '#8A8F97', marginTop: -6, marginBottom: 14 }}>
        Owed will be recalculated as ordered minus shipped.
      </div>

      <button onClick={save} style={{
        width: '100%', background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
        padding: '9px', fontSize: 13.5, fontWeight: 700, marginBottom: 8
      }}>Save changes</button>

      {confirmingDelete ? (
        <div style={{ background: '#FCEEE8', border: '1px solid #F0C4B8', borderRadius: 8, padding: 12 }}>
          <div style={{ fontSize: 12.5, color: '#B23A2E', marginBottom: 8 }}>Delete this order line? This can't be undone.</div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button onClick={() => { onDelete(order.id); onClose(); }} style={{
              flex: 1, background: '#B23A2E', color: 'white', border: 'none', borderRadius: 7, padding: '8px', fontSize: 12.5, fontWeight: 700
            }}>Delete</button>
            <button onClick={() => setConfirmingDelete(false)} style={{
              flex: 1, background: 'white', border: '1px solid #DCD9CE', borderRadius: 7, padding: '8px', fontSize: 12.5
            }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button onClick={() => setConfirmingDelete(true)} style={{
          width: '100%', background: 'white', border: '1px solid #F0C4B8', color: '#B23A2E', borderRadius: 8,
          padding: '9px', fontSize: 13, fontWeight: 700
        }}>Delete order</button>
      )}
    </ModalShell>
  );
}

function UnifiedLoginScreen({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  async function submit(e) {
    e.preventDefault();
    setError('');
    setLoading(true);
    try {
      await onLogin(username, password);
    } catch (err) {
      setError(err.message);
    }
    setLoading(false);
  }

  return (
    <div style={{ fontFamily: "'Inter', sans-serif", background: '#EEEEE9', minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@600;700&family=Inter:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        button { cursor: pointer; font-family: 'Inter', sans-serif; }
      `}</style>
      <form onSubmit={submit} style={{ background: 'white', border: '1px solid #DCD9CE', borderRadius: 12, padding: 28, width: '100%', maxWidth: 340 }}>
        <div style={{ marginBottom: 20 }}>
          <img src={LOGO_DATA_URL} alt="GR Trailers" style={{ width: '78%', height: 'auto', display: 'block', margin: '0 auto 12px' }} />
          <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 18, textTransform: 'uppercase', lineHeight: 1.1 }}>Sign In</div>
          <div style={{ fontSize: 10.5, color: '#8A8F97' }}>Staff or dealer</div>
        </div>
        <label style={{ fontSize: 11.5, fontWeight: 700, color: '#5B6470', textTransform: 'uppercase', letterSpacing: '0.03em', display: 'block', marginBottom: 5 }}>Username</label>
        <input value={username} onChange={e => setUsername(e.target.value)} style={{
          width: '100%', padding: '9px 10px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13.5, marginBottom: 12
        }} />
        <label style={{ fontSize: 11.5, fontWeight: 700, color: '#5B6470', textTransform: 'uppercase', letterSpacing: '0.03em', display: 'block', marginBottom: 5 }}>Password</label>
        <input type="password" autoFocus value={password} onChange={e => setPassword(e.target.value)} style={{
          width: '100%', padding: '9px 10px', borderRadius: 7, border: '1px solid #DCD9CE', fontSize: 13.5, marginBottom: 12
        }} />
        {error && <div style={{ color: '#B23A2E', fontSize: 12.5, marginBottom: 12 }}>{error}</div>}
        <button disabled={loading} type="submit" style={{
          width: '100%', background: '#E8592A', color: 'white', border: 'none', borderRadius: 8,
          padding: '11px', fontSize: 13.5, fontWeight: 700
        }}>{loading ? 'Signing in…' : 'Sign in'}</button>
      </form>
    </div>
  );
}

export default function App() {
  const [mode, setMode] = useState(null); // 'admin' | 'dealer' | null
  const [adminToken, setAdminTokenState] = useState('');
  const [adminSession, setAdminSessionState] = useState(null);
  const [dealerToken, setDealerTokenState] = useState('');
  const [dealerName, setDealerNameState] = useState('');
  const [checking, setChecking] = useState(true);

  useEffect(() => {
    (async () => {
      const savedAdminToken = localStorage.getItem('gr-admin-token');
      if (savedAdminToken) {
        try {
          const res = await fetch('/api/admin-auth', {
            method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'session', token: savedAdminToken }),
          });
          if (res.ok) {
            const data = await res.json();
            setAdminToken(savedAdminToken);
            setAdminTokenState(savedAdminToken);
            setAdminSessionState({ role: data.role, username: data.username, permissions: data.permissions });
            setMode('admin');
            setChecking(false);
            return;
          }
        } catch { /* fall through */ }
        localStorage.removeItem('gr-admin-token');
      }

      const savedDealerToken = localStorage.getItem('gr-dealer-token');
      if (savedDealerToken) {
        try {
          const res = await fetch('/api/dealer-auth', {
            method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'session', token: savedDealerToken }),
          });
          if (res.ok) {
            const data = await res.json();
            setDealerTokenState(savedDealerToken);
            setDealerNameState(data.dealerName);
            setMode('dealer');
            setChecking(false);
            return;
          }
        } catch { /* fall through */ }
        localStorage.removeItem('gr-dealer-token');
      }

      setChecking(false);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    onUnauthorizedRequest(() => {
      localStorage.removeItem('gr-admin-token');
      setAdminTokenState('');
      setAdminSessionState(null);
      setMode(m => (m === 'admin' ? null : m));
    });
  }, []);

  async function handleUnifiedLogin(username, password) {
    // Try staff/owner login first.
    const adminRes = await fetch('/api/admin-auth', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'login', username, password }),
    });
    if (adminRes.ok) {
      const data = await adminRes.json();
      localStorage.setItem('gr-admin-token', data.token);
      setAdminToken(data.token);
      setAdminTokenState(data.token);
      setAdminSessionState({ role: data.role, username: data.username, permissions: data.permissions });
      setMode('admin');
      return;
    }
    const adminErr = await adminRes.json().catch(() => ({}));
    if (adminRes.status === 429) throw new Error(adminErr.error || 'Too many attempts.');

    // Not a staff/owner login — try it as a dealer login instead.
    const dealerRes = await fetch('/api/dealer-auth', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'login', username, password }),
    });
    if (dealerRes.ok) {
      const data = await dealerRes.json();
      localStorage.setItem('gr-dealer-token', data.token);
      setDealerTokenState(data.token);
      setDealerNameState(data.dealerName);
      setMode('dealer');
      return;
    }
    const dealerErr = await dealerRes.json().catch(() => ({}));
    if (dealerRes.status === 429) throw new Error(dealerErr.error || 'Too many attempts.');

    throw new Error('Incorrect username or password.');
  }

  function handleAdminLogout() {
    fetch('/api/admin-auth', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'logout', token: adminToken }),
    }).catch(() => {});
    localStorage.removeItem('gr-admin-token');
    setAdminToken(null);
    setAdminTokenState('');
    setAdminSessionState(null);
    setMode(null);
  }

  function handleDealerLogout() {
    fetch('/api/dealer-auth', {
      method: 'POST', cache: 'no-store', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'logout', token: dealerToken }),
    }).catch(() => {});
    localStorage.removeItem('gr-dealer-token');
    setDealerTokenState('');
    setDealerNameState('');
    setMode(null);
  }

  if (checking) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', fontFamily: "'Inter', sans-serif", color: '#5B6470' }}>
        Loading…
      </div>
    );
  }

  if (mode === 'admin') {
    return <Dashboard onAdminLogout={handleAdminLogout} adminSession={adminSession} />;
  }

  if (mode === 'dealer') {
    return <PortalHome dealerName={dealerName} token={dealerToken} onLogout={handleDealerLogout} />;
  }

  return <UnifiedLoginScreen onLogin={handleUnifiedLogin} />;
}
