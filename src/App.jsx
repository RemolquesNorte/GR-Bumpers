import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { Package, Truck, Users, AlertTriangle, Plus, ChevronDown, ChevronRight, Search, X, Check, ArrowRight, Warehouse, Loader2, PackageCheck, PackageX, Upload, Clock, UserSearch, Edit3, PackageSearch, Factory, ClipboardList, FileSpreadsheet, TrendingDown, Lock, Trash2, KeyRound, Inbox, LogOut, DownloadCloud, UploadCloud, RotateCcw } from 'lucide-react';
import * as XLSX from 'xlsx';
import { storageGet, storageSet, setAdminToken, onUnauthorizedRequest } from './storage.js';
import { PortalHome } from './Portal.jsx';

const TRANSLATION = {"Defensa frontal completa de frente con malla (Chevrolet 03-07 Diesel)Peso 172 lbs": "FBMC 03-07", "Defensa frontal completa de frente con malla (Chevrolet 07.5-10 Diesel)Peso 172 lbs": "FBMC 7.5-10", "Defensa frontal completa de frente con malla (Chevrolet 11-14 Diesel)Peso 172 lbs": "FBMC 11-14", "Defensa frontal completa de frente con malla (Chevrolet 15-19 Diesel)Peso 172 lbs": "FBMC 15-19", "Defensa frontal completa de frente con malla (Chevrolet 20-23 Diesel) Peso 206.8 Lbs.": "FBMC 20-23", "Defensa frontal completa de frente con malla (Chevrolet 24-25 Diesel) Peso 206.8 Lbs.": "FBMC 24-25", "Defensa frontal completa de frente con malla (Chevrolet 26- Diesel) Peso 206.8 Lbs.": "FBMC 26-", "Defensa frontal completa de frente con malla (Chevrolet 99-02 Diesel)Peso 172 lbs.": "FBMC 99-02", "Defensa frontal completa de frente con malla (Dodge 03-05 Diesel)Peso 172 lbs.": "FBMD 03-05", "Defensa frontal completa de frente con malla (Dodge 06-09 Diesel) peso 172 lbs": "FBMD 06-09", "Defensa frontal completa de frente con malla (Dodge 10-18 Diesel)Peso 172 lbs": "FBMD 10-18", "Defensa frontal completa de frente con malla (Dodge 19-24 Diesel)Peso 238 lbs": "FBMD 19-24", "Defensa frontal completa de frente con malla (Dodge 95-02 Diesel)Peso 172 lbs": "FBMD 95-02", "Defensa frontal completa de frente con malla (Ford 05-07 Diesel)Peso 172 lbs": "FBMF 05-07", "Defensa frontal completa de frente con malla (Ford 08-10 Diesel)Peso 172 lbs": "FBMF 08-10", "Defensa frontal completa de frente con malla (Ford 11-16 F-350 y F-450 Diesel ) Peso 172 lbs.": "FBMF 11-16", "Defensa frontal completa de frente con malla (Ford 17-22 Diesel ) Peso 236 lbs.": "FBMF 17-22", "Defensa frontal completa de frente con malla (Ford 23-25 Diesel ) Peso 236 lbs.": "FBMF 23-26", "Defensa frontal completa de frente con malla (Ford 94-98 Diesel)": "FBMF 94-98", "Defensa frontal completa de frente con malla (Ford 99-04 Diesel)Peso 172 lbs": "FBMF 99-04", "Defensa frontal completa de frente con malla (GMC 03-07 Diesel) Peso 172 Lbs.": "FBMG 03-07", "Defensa frontal completa de frente con malla (GMC 11-14 Diesel) peso 172 lbs": "FBMG 11-14", "Defensa frontal completa de frente con malla (GMC 15-19 Diesel) Peso 172 Lbs.": "FBMG 15-19", "Defensa frontal completa de frente con malla (GMC 20-22 Diesel) Peso 172 Lbs.": "FBMG 20-22", "Defensa frontal completa de frente con malla (GMC 7.5-10 Diesel)Peso 172 lbs": "FBMG 7.5-10", "Defensa frontal completa de frente con malla (GMC 99-02 Diesel)": "FBMG 99-02", "Defensa trasera (Chevy 18-26 Diesel) peso 115 lbs.": "RBCGD 18-26", "Defensa trasera (Chevy y GMC 11-18 Diesel) peso 115 lbs.": "RBCGD 11-18", "Defensa trasera (Chevy y GMC 03-07 Diesel) Peso 115 lbs.": "RBCGD 03-07", "Defensa trasera (Chevy y GMC 07-10 Diesel) peso 115 lbs.": "RBCGD 07-10", "Defensa trasera (Chevy y GMC 20-22 Diesel) peso 115 lbs.": "RBCGD 20-22", "Defensa trasera (Chevy y GMC 99-02 Diesel)": "RBCGD 99-02", "Defensa trasera (Dodge 03-09 Diesel) Peso 115 lbs": "RBDD 03-09", "Defensa trasera (Dodge 24-26 Diesel)": "RBDD 24-26", "Defensa trasera (Dodge 95-02 Diesel) Peso 115 lbs": "RBDD 95-02", "Defensa trasera (Ford 08-16 Diesel)": "RBFD 08-16", "Defensa trasera (Ford 17-22 Diesel)": "RBFD 17-22", "Defensa trasera (Ford 23-26 Diesel)": "RBFD 23-26", "Defensa trasera (Ford 94-98 Diesel)": "RBFD 94-98", "Defensa trasera (Ford 99-07 Diesel) Peso 115 lbs": "RBFD 99-07", "Defensa trasera con luces (Chevy y GMC 99-02 Diesel)": "RBCGD 99-02"}
;
const SEED = {"inventory": [{"sku": "FBMC 03-07", "mx": 15, "us": 33}, {"sku": "FBMC 11-14", "mx": 13, "us": 0}, {"sku": "FBMC 15-19", "mx": 42, "us": 0}, {"sku": "FBMC 20-23", "mx": 26, "us": 35}, {"sku": "FBMC 24-25", "mx": 39, "us": 0}, {"sku": "FBMC 24-26", "mx": 0, "us": 63}, {"sku": "FBMC 26-", "mx": 1, "us": 0}, {"sku": "FBMC 7.5-10", "mx": 20, "us": 8}, {"sku": "FBMC 99-02", "mx": 15, "us": 0}, {"sku": "FBMD 03-05", "mx": 10, "us": 19}, {"sku": "FBMD 06-09", "mx": 25, "us": 0}, {"sku": "FBMD 10-18", "mx": 82, "us": 7}, {"sku": "FBMD 19-24", "mx": 95, "us": 15}, {"sku": "FBMD 95-02", "mx": 24, "us": 6}, {"sku": "FBMF 05-07", "mx": 18, "us": 9}, {"sku": "FBMF 08-10", "mx": 38, "us": 35}, {"sku": "FBMF 11-16", "mx": 63, "us": 4}, {"sku": "FBMF 17-22", "mx": 121, "us": 8}, {"sku": "FBMF 23-25", "mx": 86, "us": 0}, {"sku": "FBMF 94-98", "mx": 36, "us": 0}, {"sku": "FBMF 99-04", "mx": 22, "us": 0}, {"sku": "FBMG 03-07", "mx": 14, "us": 9}, {"sku": "FBMG 11-14", "mx": 18, "us": 0}, {"sku": "FBMG 15-19", "mx": 10, "us": 51}, {"sku": "FBMG 20-22", "mx": 39, "us": 0}, {"sku": "FBMG 20-23", "mx": 0, "us": 24}, {"sku": "FBMG 7.5-10", "mx": 13, "us": 50}, {"sku": "FBMG 99-02", "mx": 13, "us": 7}, {"sku": "RBCGD 03-07", "mx": 5, "us": 0}, {"sku": "RBCGD 07-10", "mx": 18, "us": 0}, {"sku": "RBCGD 11-18", "mx": 13, "us": 0}, {"sku": "RBCGD 18-26", "mx": 1, "us": 0}, {"sku": "RBCGD 20-22", "mx": 9, "us": 0}, {"sku": "RBCGD 20-23", "mx": 0, "us": 16}, {"sku": "RBCGD 99-02", "mx": 23, "us": 6}, {"sku": "RBDD 03-09", "mx": 27, "us": 10}, {"sku": "RBDD 24-26", "mx": 1, "us": 0}, {"sku": "RBDD 95-02", "mx": 24, "us": 21}, {"sku": "RBFD 08-16", "mx": 8, "us": 0}, {"sku": "RBFD 17-22", "mx": 13, "us": 5}, {"sku": "RBFD 23-26", "mx": 2, "us": 0}, {"sku": "RBFD 94-98", "mx": 6, "us": 23}, {"sku": "RBFD 99-07", "mx": 21, "us": 7}], "toolboxItems": [{"item_es": "Caja de Herramientas de 18\"ancho  x 12\" fondo x 18\" alto", "sku": null, "qty": 10}, {"item_es": "Caja de Herramientas de 24\" ancho  x 12\" fondo x 18\"alto", "sku": null, "qty": 22}, {"item_es": "Caja de Herramientas de 30\" largo  x 20\" ancho x 16\"alto", "sku": null, "qty": 10}], "dealers": [{"name": "2-D Trailer Sales LLC", "origin": "US"}, {"name": "4 Mile Trailers", "origin": "US"}, {"name": "4-W Cattle Company, LLC", "origin": "US"}, {"name": "A Day Trailers, LLC", "origin": "MX"}, {"name": "APC Equipment, Inc.", "origin": "MX"}, {"name": "Accel logistics", "origin": "US"}, {"name": "Ag 98 Trailer Sales", "origin": "MX"}, {"name": "Atkins Richard", "origin": "US"}, {"name": "BT Tire & Ag Sales, LLC", "origin": "MX"}, {"name": "Bailey Auto", "origin": "MX"}, {"name": "Bar Circle W. Sales, LLC", "origin": "US"}, {"name": "Bartlett Cooperative Assn.", "origin": "US"}, {"name": "Battleborn Trailer Sales, LLC", "origin": "MX"}, {"name": "Bell Trailerplex", "origin": "MX"}, {"name": "Better Built Trailers", "origin": "MX"}, {"name": "C & J Traders", "origin": "MX"}, {"name": "California Custom Trailers", "origin": "MX"}, {"name": "Christensen Ranches", "origin": "MX"}, {"name": "Clayton Brothers Farm Trust", "origin": "MX"}, {"name": "Colt Bruegman Trailer Sales, LLC", "origin": "MX"}, {"name": "Cross Trail", "origin": "MX"}, {"name": "Crouch Mesa Trailer Sales, LLC", "origin": "MX"}, {"name": "D & E Sales", "origin": "US"}, {"name": "DH Farm Equipment", "origin": "US"}, {"name": "DP Platinum Star Trailers, LLC", "origin": "US"}, {"name": "Diamond K Sales", "origin": "MX"}, {"name": "Diamond T Metals", "origin": "US"}, {"name": "Diamond T Sales", "origin": "US"}, {"name": "Don Evans Legal Window Tint & Trailers", "origin": "US"}, {"name": "Double S Towing Service, LLC", "origin": "US"}, {"name": "Double Z Service & Supply, LLC", "origin": "MX"}, {"name": "Durham Trailer Ranch, Inc.", "origin": "MX"}, {"name": "Early Trailer Sales, LLC", "origin": "MX"}, {"name": "Elite Trailer Sales & Service, LLC", "origin": "US"}, {"name": "F & S Take Off Parts", "origin": "US"}, {"name": "Frye Farms Trailers", "origin": "MX"}, {"name": "GR Trailers LLC. {Customer}", "origin": "MX"}, {"name": "Gengler Auto, LLC", "origin": "US"}, {"name": "Goodman Ag Supply, Inc.", "origin": "MX"}, {"name": "H 5 Ranch, LLC", "origin": "US"}, {"name": "Heacock Trailers & Truck", "origin": "MX"}, {"name": "Hendrys Ron", "origin": "US"}, {"name": "Hisle Brothers, Inc.", "origin": "US"}, {"name": "Hitchin Post Motors", "origin": "US"}, {"name": "Horsch Trailer Sales", "origin": "US"}, {"name": "J&G Trailers", "origin": "MX"}, {"name": "JME Trailers, LLC", "origin": "MX"}, {"name": "Jason Lewis Automotive", "origin": "MX"}, {"name": "Jim's Motors INC", "origin": "MX"}, {"name": "Ken's Trailer Sales & Repair", "origin": "US"}, {"name": "King Gary", "origin": "MX"}, {"name": "Kohler Trailer Sales", "origin": "MX"}, {"name": "Legion Diesel's LLC", "origin": "US"}, {"name": "Lone Star Trailers", "origin": "MX"}, {"name": "MJ Trailers, LLC", "origin": "US"}, {"name": "McFadden Trailer Sales", "origin": "MX"}, {"name": "McMillan AG Repair & Service", "origin": "US"}, {"name": "Meyer Automotive", "origin": "MX"}, {"name": "Mid-Valley Trailer Sales", "origin": "MX"}, {"name": "Monday Trailers & Equipment", "origin": "US"}, {"name": "Northwest Farm Supply", "origin": "MX"}, {"name": "Open Range, LLC", "origin": "US"}, {"name": "Penner Trailer Sales, LLC", "origin": "MX"}, {"name": "Pinnacle Trailer Sales, LLC", "origin": "MX"}, {"name": "Prime Rate Motors, INC.", "origin": "MX"}, {"name": "Producers CO-OP Trailers", "origin": "MX"}, {"name": "Red Barn Trailers, LLC", "origin": "MX"}, {"name": "Redline Trailers and More", "origin": "US"}, {"name": "Riverside Boot & Saddle Blackfoot", "origin": "MX"}, {"name": "Riverside Trailers Caldwell", "origin": "MX"}, {"name": "Riverside Trailers Jerome", "origin": "MX"}, {"name": "Rod's Auto Sales LLC", "origin": "US"}, {"name": "Route 66 Trailer Sales", "origin": "US"}, {"name": "Royer John", "origin": "US"}, {"name": "S & S Motors LLC", "origin": "US"}, {"name": "Scranton Truck & Trailer", "origin": "MX"}, {"name": "Shalom Trailers, INC", "origin": "MX"}, {"name": "Springfield Trailers, INC.", "origin": "US"}, {"name": "Superior Steel Sales, LLC", "origin": "US"}, {"name": "Superior Trailer Sales", "origin": "US"}, {"name": "T & T Quality Building", "origin": "US"}, {"name": "T & T Trailer Sales LLC", "origin": "MX"}, {"name": "TIP TOP TRAILERS", "origin": "MX"}, {"name": "Temco Mfg. Inc.", "origin": "MX"}, {"name": "The Part Xperts", "origin": "MX"}, {"name": "The Truck Shop, LLC", "origin": "US"}, {"name": "Timberline Truck & Trailer", "origin": "US"}, {"name": "Trail Dust Trailers INC", "origin": "MX"}, {"name": "Traileros", "origin": "MX"}, {"name": "Trailquip Plus, LLC", "origin": "US"}, {"name": "Travln Toys", "origin": "MX"}, {"name": "Tri-Star Fleet Sales", "origin": "MX"}, {"name": "Triple M Trailers", "origin": "MX"}, {"name": "Tru-Trailers, Inc.", "origin": "MX"}, {"name": "Truck Country", "origin": "MX"}, {"name": "Unlimited Genetics, Inc.", "origin": "US"}, {"name": "V-Bar Trailer Sales", "origin": "MX"}, {"name": "Walker Trailer And Equipment, LLC", "origin": "US"}, {"name": "West Luke", "origin": "US"}, {"name": "West Texas Trailers & Equipment", "origin": "US"}, {"name": "Wilcoxson Welding", "origin": "US"}, {"name": "Wolfe Auto Service, LLC", "origin": "US"}, {"name": "XB Trailer Sales", "origin": "MX"}, {"name": "YBell Ranch Supply", "origin": "MX"}, {"name": "Zona Trailer Sales, LLC", "origin": "MX"}], "orders": [{"id": "O1", "sku": "BM-FBMC 03-07C", "date": "06/19/2026", "due": "06/19/2026", "num": "39369", "po": "", "dealer": "Cash", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O2", "sku": "FBMC 03-07", "date": "02/27/2026", "due": "03/29/2026", "num": "37987", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O3", "sku": "FBMC 03-07", "date": "07/03/2026", "due": "06/05/2026", "num": "38078", "po": "", "dealer": "Legion Diesel's LLC", "qty": 8, "invoiced": 7, "backordered": 1}, {"id": "O4", "sku": "FBMC 03-07", "date": "04/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 5, "backordered": 5}, {"id": "O5", "sku": "FBMC 03-07", "date": "04/20/2026", "due": "05/20/2026", "num": "38614", "po": "", "dealer": "Christensen Ranches", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O6", "sku": "FBMC 03-07", "date": "12/05/2026", "due": "11/06/2026", "num": "38935", "po": "", "dealer": "Horsch Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O7", "sku": "FBMC 03-07", "date": "5/29/2026", "due": "06/28/2026", "num": "39120", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O8", "sku": "FBMC 03-07", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O9", "sku": "FBMC 11-14", "date": "02/06/2026", "due": "02/07/2026", "num": "39145", "po": "", "dealer": "S & S Motors LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O10", "sku": "FBMC 11-14", "date": "08/06/2026", "due": "08/07/2026", "num": "39209", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O11", "sku": "FBMC 11-14", "date": "09/06/2026", "due": "09/07/2026", "num": "39239", "po": "", "dealer": "Trailquip Plus, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O12", "sku": "FBMC 11-14", "date": "6/17/2026", "due": "7/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O13", "sku": "FBMC 15-19", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O14", "sku": "FBMC 15-19", "date": "06/02/2026", "due": "10/03/2026", "num": "37729", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O15", "sku": "FBMC 15-19", "date": "04/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 9, "backordered": 1}, {"id": "O16", "sku": "FBMC 15-19", "date": "04/14/2026", "due": "05/14/2026", "num": "38529", "po": "", "dealer": "Open Range, LLC", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O17", "sku": "FBMC 15-19", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O18", "sku": "FBMC 15-19", "date": "6/18/2026", "due": "07/18/2026", "num": "39348", "po": "21257", "dealer": "Temco Mfg. Inc.", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O19", "sku": "FBMC 15-19", "date": "6/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O20", "sku": "FBMC 20-23", "date": "12/06/2025", "due": "12/07/2025", "num": "34882", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 9, "backordered": 1}, {"id": "O21", "sku": "FBMC 20-23", "date": "08/15/2025", "due": "09/14/2025", "num": "35640", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O22", "sku": "FBMC 20-23", "date": "06/01/2026", "due": "05/02/2026", "num": "37279", "po": "Dean", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O23", "sku": "FBMC 20-23", "date": "07/03/2026", "due": "06/05/2026", "num": "38078", "po": "", "dealer": "Legion Diesel's LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O24", "sku": "FBMC 20-23", "date": "01/07/2026", "due": "10/08/2026", "num": "39495", "po": "", "dealer": "BT Tire & Ag Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O25", "sku": "FBMC 24-26", "date": "12/06/2025", "due": "12/07/2025", "num": "34882", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 0, "backordered": 10}, {"id": "O26", "sku": "FBMC 24-26", "date": "03/11/2025", "due": "03/12/2025", "num": "36637", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O27", "sku": "FBMC 24-26", "date": "11/03/2026", "due": "10/04/2026", "num": "38112", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O28", "sku": "FBMC 24-26", "date": "11/03/2026", "due": "10/04/2026", "num": "38116", "po": "", "dealer": "JME Trailers, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O29", "sku": "FBMC 24-26", "date": "4/20/2026", "due": "5/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O30", "sku": "FBMC 24-26", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O31", "sku": "FBMC 24-26", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O32", "sku": "FBMC 24-26", "date": "09/06/2026", "due": "09/07/2026", "num": "39244", "po": "", "dealer": "Kohler Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O33", "sku": "FBMC 24-26", "date": "6/24/2026", "due": "07/24/2026", "num": "39419", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O34", "sku": "FBMC 7.5-10", "date": "11/11/2025", "due": "10/12/2025", "num": "36752", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O35", "sku": "FBMC 7.5-10", "date": "07/01/2026", "due": "06/02/2026", "num": "37289", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 7, "backordered": 3}, {"id": "O36", "sku": "FBMC 7.5-10", "date": "3/31/2026", "due": "04/30/2026", "num": "38376", "po": "", "dealer": "Gengler Auto, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O37", "sku": "FBMC 7.5-10", "date": "4/20/2026", "due": "05/20/2026", "num": "38618", "po": "21094", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O38", "sku": "FBMC 7.5-10", "date": "04/22/2026", "due": "5/22/2026", "num": "38690", "po": "21102", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O39", "sku": "FBMC 7.5-10", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O40", "sku": "FBMC 7.5-10", "date": "6/29/2026", "due": "7/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O41", "sku": "FBMC 99-02", "date": "09/06/2026", "due": "09/07/2026", "num": "39234", "po": "", "dealer": "West Luke", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O42", "sku": "FBMC 99-02", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O43", "sku": "FBMD 03-05", "date": "1/26/2026", "due": "2/25/2026", "num": "37560", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 3, "backordered": 2}, {"id": "O44", "sku": "FBMD 03-05", "date": "08/06/2026", "due": "08/07/2026", "num": "39216", "po": "", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O45", "sku": "FBMD 03-05", "date": "12/06/2026", "due": "10/07/2026", "num": "39288", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O46", "sku": "FBMD 03-05", "date": "6/29/2026", "due": "7/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O47", "sku": "FBMD 03-05", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O48", "sku": "FBMD 06-09", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O49", "sku": "FBMD 06-09", "date": "03/16/2026", "due": "10/04/2026", "num": "38194", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O50", "sku": "FBMD 06-09", "date": "03/18/2026", "due": "04/17/2026", "num": "38232", "po": "13025", "dealer": "Diamond K Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O51", "sku": "FBMD 06-09", "date": "03/31/2026", "due": "04/30/2026", "num": "38372", "po": "", "dealer": "S & S Motors LLC", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O52", "sku": "FBMD 06-09", "date": "04/14/2026", "due": "05/14/2026", "num": "38545", "po": "", "dealer": "West Luke", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O53", "sku": "FBMD 06-09", "date": "12/06/2026", "due": "12/07/2026", "num": "39289", "po": "", "dealer": "S & S Motors LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O54", "sku": "FBMD 06-09", "date": "06/23/2026", "due": "07/23/2026", "num": "39406", "po": "", "dealer": "Lone Star Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O55", "sku": "FBMD 06-09", "date": "06/30/2026", "due": "7/30/2026", "num": "39467", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O56", "sku": "FBMD 10-18", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O57", "sku": "FBMD 10-18", "date": "03/25/2026", "due": "4/24/2026", "num": "38329", "po": "21037", "dealer": "Temco Mfg. Inc.", "qty": 4, "invoiced": 3, "backordered": 1}, {"id": "O58", "sku": "FBMD 10-18", "date": "06/04/2026", "due": "06/05/2026", "num": "38432", "po": "", "dealer": "S & S Motors LLC", "qty": 10, "invoiced": 7, "backordered": 3}, {"id": "O59", "sku": "FBMD 10-18", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O60", "sku": "FBMD 10-18", "date": "08/06/2026", "due": "08/07/2026", "num": "39215", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O61", "sku": "FBMD 10-18", "date": "6/17/2026", "due": "7/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O62", "sku": "FBMD 10-18", "date": "6/17/2026", "due": "10/07/2026", "num": "39330", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O63", "sku": "FBMD 10-18", "date": "6/22/2026", "due": "06/22/2026", "num": "39374", "po": "", "dealer": "King Gary", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O64", "sku": "FBMD 10-18", "date": "06/22/2026", "due": "10/07/2026", "num": "39379", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O65", "sku": "FBMD 10-18", "date": "06/29/2026", "due": "07/29/2026", "num": "39459", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O66", "sku": "FBMD 10-18", "date": "6/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O67", "sku": "FBMD 10-18", "date": "01/07/2026", "due": "07/31/2026", "num": "39486", "po": "", "dealer": "Early Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O68", "sku": "FBMD 10-18", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O69", "sku": "FBMD 19-24", "date": "12/30/2025", "due": "01/29/2026", "num": "37207", "po": "", "dealer": "Legion Diesel's LLC", "qty": 15, "invoiced": 10, "backordered": 5}, {"id": "O70", "sku": "FBMD 19-24", "date": "07/01/2026", "due": "06/02/2026", "num": "37289", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 5, "backordered": 5}, {"id": "O71", "sku": "FBMD 19-24", "date": "04/03/2026", "due": "03/04/2026", "num": "38045", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O72", "sku": "FBMD 19-24", "date": "06/03/2026", "due": "05/04/2026", "num": "38073", "po": "3528", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O73", "sku": "FBMD 19-24", "date": "03/17/2026", "due": "4/16/2026", "num": "38209", "po": "3543", "dealer": "Truck Country", "qty": 3, "invoiced": 1, "backordered": 2}, {"id": "O74", "sku": "FBMD 19-24", "date": "03/31/2026", "due": "4/30/2026", "num": "38377", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O75", "sku": "FBMD 19-24", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O76", "sku": "FBMD 19-24", "date": "10/06/2026", "due": "10/07/2026", "num": "39258", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O77", "sku": "FBMD 19-24", "date": "10/06/2026", "due": "10/07/2026", "num": "39258", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O78", "sku": "FBMD 19-24", "date": "06/17/2026", "due": "07/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 6, "invoiced": 2, "backordered": 4}, {"id": "O79", "sku": "FBMD 19-24", "date": "06/22/2026", "due": "06/22/2026", "num": "39374", "po": "", "dealer": "King Gary", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O80", "sku": "FBMD 19-24", "date": "06/24/2026", "due": "07/24/2026", "num": "39419", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O81", "sku": "FBMD 19-24", "date": "06/25/2026", "due": "10/07/2026", "num": "39428", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O82", "sku": "FBMD 19-24", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O83", "sku": "FBMD 25-26", "date": "02/12/2025", "due": "01/01/2026", "num": "36978", "po": "", "dealer": "S & S Motors LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O84", "sku": "FBMD 25-26", "date": "12/23/2025", "due": "01/22/2026", "num": "37186", "po": "", "dealer": "Diamond T Metals", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O85", "sku": "FBMD 25-26", "date": "02/01/2026", "due": "01/02/2026", "num": "37240", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O86", "sku": "FBMD 25-26", "date": "02/02/2026", "due": "04/03/2026", "num": "37637", "po": "3471", "dealer": "Truck Country", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O87", "sku": "FBMD 25-26", "date": "03/02/2026", "due": "05/03/2026", "num": "37655", "po": "", "dealer": "Kohler Trailer Sales", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O88", "sku": "FBMD 25-26", "date": "05/02/2026", "due": "07/03/2026", "num": "37707", "po": "", "dealer": "Route 66 Trailer Sales", "qty": 8, "invoiced": 0, "backordered": 8}, {"id": "O89", "sku": "FBMD 25-26", "date": "12/02/2026", "due": "03/14/2026", "num": "37802", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O90", "sku": "FBMD 25-26", "date": "06/03/2026", "due": "05/04/2026", "num": "38074", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O91", "sku": "FBMD 25-26", "date": "12/03/2026", "due": "11/04/2026", "num": "38128", "po": "", "dealer": "Texoma Trailers, Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O92", "sku": "FBMD 25-26", "date": "3/23/2026", "due": "07/04/2026", "num": "38296", "po": "", "dealer": "XB Trailer Sales", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O93", "sku": "FBMD 25-26", "date": "4/14/2026", "due": "5/14/2026", "num": "38533", "po": "15-5564", "dealer": "Trailquip Plus, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O94", "sku": "FBMD 25-26", "date": "4/20/2026", "due": "5/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O95", "sku": "FBMD 25-26", "date": "4/30/2026", "due": "05/30/2026", "num": "38817", "po": "", "dealer": "Hitchin Post Motors", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O96", "sku": "FBMD 25-26", "date": "05/05/2026", "due": "04/06/2026", "num": "38859", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O97", "sku": "FBMD 25-26", "date": "11/05/2026", "due": "10/06/2026", "num": "38931", "po": "", "dealer": "D & E Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O98", "sku": "FBMD 25-26", "date": "6/17/2026", "due": "7/17/2026", "num": "39328", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 32, "invoiced": 0, "backordered": 32}, {"id": "O99", "sku": "FBMD 25-26", "date": "6/18/2026", "due": "07/18/2026", "num": "39347", "po": "", "dealer": "Lone Star Trailers", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O100", "sku": "FBMD 25-26", "date": "6/18/2026", "due": "7/18/2026", "num": "39348", "po": "21257", "dealer": "Temco Mfg. Inc.", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O101", "sku": "FBMD 25-26", "date": "6/19/2026", "due": "7/19/2026", "num": "39359", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O102", "sku": "FBMD 25-26", "date": "6/19/2026", "due": "7/19/2026", "num": "39360", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 10, "invoiced": 0, "backordered": 10}, {"id": "O103", "sku": "FBMD 25-26", "date": "6/19/2026", "due": "7/19/2026", "num": "39361", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 0, "backordered": 10}, {"id": "O104", "sku": "FBMD 25-26", "date": "06/19/2026", "due": "07/19/2026", "num": "39362", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O105", "sku": "FBMD 25-26", "date": "06/19/2026", "due": "07/19/2026", "num": "39363", "po": "", "dealer": "D & E Sales", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O106", "sku": "FBMD 25-26", "date": "06/19/2026", "due": "06/19/2026", "num": "39365", "po": "", "dealer": "RZS, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O107", "sku": "FBMD 25-26", "date": "06/22/2026", "due": "07/22/2026", "num": "39371", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O108", "sku": "FBMD 25-26", "date": "06/22/2026", "due": "10/07/2026", "num": "39377", "po": "", "dealer": "JME Trailers, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O109", "sku": "FBMD 25-26", "date": "06/22/2026", "due": "10/07/2026", "num": "39388", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O110", "sku": "FBMD 25-26", "date": "6/23/2026", "due": "6/23/2026", "num": "39398", "po": "", "dealer": "King Gary", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O111", "sku": "FBMD 25-26", "date": "06/25/2026", "due": "7/25/2026", "num": "39433", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O112", "sku": "FBMD 25-26", "date": "06/25/2026", "due": "07/25/2026", "num": "39434", "po": "", "dealer": "Bar Circle W. Sales of Texas, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O113", "sku": "FBMD 95-02", "date": "4/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O114", "sku": "FBMD 95-02", "date": "4/15/2026", "due": "4/15/2026", "num": "38556", "po": "", "dealer": "Route 66 Trailer Sales", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O115", "sku": "FBMD 95-02", "date": "12/05/2026", "due": "11/06/2026", "num": "38934", "po": "21148", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O116", "sku": "FBMD 95-02", "date": "5/27/2026", "due": "06/26/2026", "num": "39087", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O117", "sku": "FBMD 95-02", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O118", "sku": "FBMF 05-07", "date": "11/24/2025", "due": "12/24/2025", "num": "36901", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O119", "sku": "FBMF 05-07", "date": "12/18/2025", "due": "01/17/2026", "num": "37170", "po": "", "dealer": "Gengler Auto, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O120", "sku": "FBMF 05-07", "date": "12/05/2026", "due": "11/06/2026", "num": "38935", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O121", "sku": "FBMF 08-10", "date": "07/03/2026", "due": "06/05/2026", "num": "38078", "po": "", "dealer": "Legion Diesel's LLC", "qty": 5, "invoiced": 4, "backordered": 1}, {"id": "O122", "sku": "FBMF 11-16", "date": "1/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O123", "sku": "FBMF 11-16", "date": "3/14/2026", "due": "04/13/2026", "num": "38179", "po": "13025", "dealer": "Diamond K Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O124", "sku": "FBMF 11-16", "date": "3/20/2026", "due": "10/04/2026", "num": "38272", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O125", "sku": "FBMF 11-16", "date": "06/04/2026", "due": "06/05/2026", "num": "38412", "po": "", "dealer": "JAR Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O126", "sku": "FBMF 11-16", "date": "4/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 2, "backordered": 8}, {"id": "O127", "sku": "FBMF 11-16", "date": "05/22/2026", "due": "06/21/2026", "num": "39054", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O128", "sku": "FBMF 11-16", "date": "02/06/2026", "due": "02/07/2026", "num": "39137", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O129", "sku": "FBMF 11-16", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O130", "sku": "FBMF 11-16", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O131", "sku": "FBMF 11-16", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O132", "sku": "FBMF 11-16", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O133", "sku": "FBMF 17-22", "date": "07/15/2025", "due": "08/14/2025", "num": "35211", "po": "", "dealer": "The Part Xperts", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O134", "sku": "FBMF 17-22", "date": "01/28/2026", "due": "02/27/2026", "num": "37584", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O135", "sku": "FBMF 17-22", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 6, "invoiced": 1, "backordered": 5}, {"id": "O136", "sku": "FBMF 17-22", "date": "02/02/2026", "due": "04/03/2026", "num": "37640", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O137", "sku": "FBMF 17-22", "date": "02/16/2026", "due": "03/18/2026", "num": "37830", "po": "", "dealer": "Christensen Ranches", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O138", "sku": "FBMF 17-22", "date": "03/16/2026", "due": "04/15/2026", "num": "38183", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O139", "sku": "FBMF 17-22", "date": "3/30/2026", "due": "10/04/2026", "num": "38371", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O140", "sku": "FBMF 17-22", "date": "4/20/2026", "due": "05/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O141", "sku": "FBMF 17-22", "date": "02/06/2026", "due": "02/07/2026", "num": "39137", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 8, "invoiced": 0, "backordered": 8}, {"id": "O142", "sku": "FBMF 17-22", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O143", "sku": "FBMF 17-22", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 6, "invoiced": 0, "backordered": 6}, {"id": "O144", "sku": "FBMF 17-22", "date": "06/17/2026", "due": "07/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 7, "invoiced": 6, "backordered": 1}, {"id": "O145", "sku": "FBMF 17-22", "date": "06/25/2026", "due": "07/25/2026", "num": "39433", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O146", "sku": "FBMF 17-22", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O147", "sku": "FBMF 23-26", "date": "11/25/2025", "due": "12/25/2025", "num": "36937", "po": "", "dealer": "Legion Diesel's LLC", "qty": 10, "invoiced": 5, "backordered": 5}, {"id": "O148", "sku": "FBMF 23-26", "date": "07/01/2026", "due": "06/02/2026", "num": "37289", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 1, "backordered": 9}, {"id": "O149", "sku": "FBMF 23-26", "date": "11/03/2026", "due": "10/04/2026", "num": "38112", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 5, "invoiced": 2, "backordered": 3}, {"id": "O150", "sku": "FBMF 23-26", "date": "03/31/2026", "due": "04/30/2026", "num": "38377", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O151", "sku": "FBMF 23-26", "date": "04/13/2026", "due": "5/13/2026", "num": "38507", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O152", "sku": "FBMF 23-26", "date": "04/15/2026", "due": "5/15/2026", "num": "38555", "po": "", "dealer": "Horsch Trailer Sales", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O153", "sku": "FBMF 23-26", "date": "04/20/2026", "due": "5/20/2026", "num": "38613", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O154", "sku": "FBMF 23-26", "date": "12/05/2026", "due": "11/06/2026", "num": "38943", "po": "21149", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O155", "sku": "FBMF 23-26", "date": "05/18/2026", "due": "06/17/2026", "num": "39001", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 20, "invoiced": 0, "backordered": 20}, {"id": "O156", "sku": "FBMF 23-26", "date": "5/29/2026", "due": "6/28/2026", "num": "39120", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O157", "sku": "FBMF 23-26", "date": "02/06/2026", "due": "02/07/2026", "num": "39138", "po": "", "dealer": "Clayton Brothers Farm Trust", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O158", "sku": "FBMF 23-26", "date": "10/06/2026", "due": "10/07/2026", "num": "39255", "po": "", "dealer": "S & S Motors LLC", "qty": 7, "invoiced": 3, "backordered": 4}, {"id": "O159", "sku": "FBMF 23-26", "date": "6/17/2026", "due": "07/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O160", "sku": "FBMF 23-26", "date": "6/25/2026", "due": "07/25/2026", "num": "39433", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O161", "sku": "FBMF 23-26", "date": "6/29/2026", "due": "07/29/2026", "num": "39453", "po": "1630", "dealer": "Diamond T Metals", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O162", "sku": "FBMF 23-26", "date": "06/30/2026", "due": "10/07/2026", "num": "39473", "po": "", "dealer": "Hitchin Post Motors", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O163", "sku": "FBMF 23-26", "date": "03/07/2026", "due": "10/08/2026", "num": "39523", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O164", "sku": "FBMF 94-98", "date": "03/25/2026", "due": "04/24/2026", "num": "38329", "po": "21037", "dealer": "Temco Mfg. Inc.", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O165", "sku": "FBMF 94-98", "date": "04/29/2026", "due": "10/05/2026", "num": "38792", "po": "", "dealer": "DH Farm Equipment", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O166", "sku": "FBMF 94-98", "date": "06/17/2026", "due": "7/17/2026", "num": "39326", "po": "1595", "dealer": "Crouch Mesa Trailer Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O167", "sku": "FBMF 94-98", "date": "01/07/2026", "due": "7/31/2026", "num": "39487", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O168", "sku": "FBMF 99-04", "date": "11/11/2025", "due": "10/12/2025", "num": "36752", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O169", "sku": "FBMF 99-04", "date": "01/16/2026", "due": "2/15/2026", "num": "37415", "po": "", "dealer": "S & S Motors LLC", "qty": 7, "invoiced": 5, "backordered": 2}, {"id": "O170", "sku": "FBMF 99-04", "date": "02/27/2026", "due": "03/29/2026", "num": "37987", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 3, "invoiced": 2, "backordered": 1}, {"id": "O171", "sku": "FBMF 99-04", "date": "03/13/2026", "due": "12/04/2026", "num": "38155", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 10, "invoiced": 9, "backordered": 1}, {"id": "O172", "sku": "FBMF 99-04", "date": "12/05/2026", "due": "11/06/2026", "num": "38935", "po": "", "dealer": "Horsch Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O173", "sku": "FBMF 99-04", "date": "05/22/2026", "due": "06/21/2026", "num": "39054", "po": "", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O174", "sku": "FBMF 99-04", "date": "10/06/2026", "due": "10/07/2026", "num": "39250", "po": "", "dealer": "Better Built Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O175", "sku": "FBMF 99-04", "date": "06/15/2026", "due": "07/15/2026", "num": "39304", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 5, "invoiced": 0, "backordered": 5}, {"id": "O176", "sku": "FBMF 99-04", "date": "06/29/2026", "due": "07/29/2026", "num": "39459", "po": "", "dealer": "Battleborn Trailer Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O177", "sku": "FBMF 99-04", "date": "01/07/2026", "due": "07/31/2026", "num": "39501", "po": "Abegglen", "dealer": "Ken's Trailer Sales & Repair, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O178", "sku": "FBMG 03-07", "date": "12/02/2026", "due": "03/14/2026", "num": "37796", "po": "", "dealer": "XB Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O179", "sku": "FBMG 03-07", "date": "12/02/2026", "due": "03/14/2026", "num": "37802", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O180", "sku": "FBMG 03-07", "date": "02/16/2026", "due": "03/18/2026", "num": "37844", "po": "", "dealer": "XB Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O181", "sku": "FBMG 03-07", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O182", "sku": "FBMG 03-07", "date": "02/07/2026", "due": "10/08/2026", "num": "39513", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O183", "sku": "FBMG 11-14", "date": "06/03/2026", "due": "05/05/2026", "num": "38060", "po": "", "dealer": "Legion Diesel's LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O184", "sku": "FBMG 11-14", "date": "11/04/2026", "due": "11/05/2026", "num": "38495", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O185", "sku": "FBMG 11-14", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O186", "sku": "FBMG 11-14", "date": "06/22/2026", "due": "6/22/2026", "num": "39380", "po": "", "dealer": "Cash", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O187", "sku": "FBMG 15-19", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O188", "sku": "FBMG 15-19", "date": "04/27/2026", "due": "5/27/2026", "num": "38752", "po": "", "dealer": "S & S Motors LLC", "qty": 2, "invoiced": 1, "backordered": 1}, {"id": "O189", "sku": "FBMG 15-19", "date": "03/07/2026", "due": "02/08/2026", "num": "39524", "po": "", "dealer": "Gengler Auto, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O190", "sku": "FBMG 20-23", "date": "01/30/2026", "due": "01/03/2026", "num": "37622", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O191", "sku": "FBMG 20-23", "date": "03/07/2026", "due": "10/08/2026", "num": "39522", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O192", "sku": "FBMG 24-26", "date": "08/08/2023", "due": "07/09/2023", "num": "27566", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O193", "sku": "FBMG 24-26", "date": "02/02/2024", "due": "03/03/2024", "num": "29440", "po": "", "dealer": "Wheeler Bar Circle W. Sales, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O194", "sku": "FBMG 24-26", "date": "02/28/2024", "due": "2/28/2024", "num": "29722", "po": "", "dealer": "Durham Trailer Ranch, Inc.", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O195", "sku": "FBMG 24-26", "date": "2/29/2024", "due": "2/29/2024", "num": "29745", "po": "", "dealer": "Route 66 Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O196", "sku": "FBMG 24-26", "date": "3/13/2024", "due": "12/04/2024", "num": "29909", "po": "", "dealer": "XB Trailer Sales", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O197", "sku": "FBMG 24-26", "date": "02/04/2024", "due": "02/04/2024", "num": "30130", "po": "", "dealer": "Wolfe Auto Service, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O198", "sku": "FBMG 24-26", "date": "05/06/2024", "due": "05/06/2024", "num": "30898", "po": "", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O199", "sku": "FBMG 24-26", "date": "11/07/2024", "due": "11/07/2024", "num": "31303", "po": "", "dealer": "West Luke", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O200", "sku": "FBMG 24-26", "date": "7/25/2024", "due": "8/24/2024", "num": "31466", "po": "", "dealer": "Superior Steel Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O201", "sku": "FBMG 24-26", "date": "8/22/2024", "due": "9/21/2024", "num": "32635", "po": "", "dealer": "Legion Diesel's LLC", "qty": 4, "invoiced": 0, "backordered": 4}, {"id": "O202", "sku": "FBMG 24-26", "date": "12/11/2024", "due": "12/12/2024", "num": "32586", "po": "", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O203", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32662", "po": "Shawn", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O204", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32663", "po": "Donald", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O205", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32664", "po": "Amanda", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O206", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32665", "po": "Was", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O207", "sku": "FBMG 24-26", "date": "11/19/2024", "due": "12/19/2024", "num": "32666", "po": "Kyle", "dealer": "Route 66 Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O208", "sku": "FBMG 24-26", "date": "12/13/2024", "due": "12/01/2025", "num": "32915", "po": "", "dealer": "Christensen Ranches", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O209", "sku": "FBMG 24-26", "date": "12/16/2024", "due": "01/15/2025", "num": "32924", "po": "", "dealer": "Diamond K Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O210", "sku": "FBMG 24-26", "date": "1/27/2025", "due": "2/26/2025", "num": "33288", "po": "", "dealer": "Diamond K Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O211", "sku": "FBMG 24-26", "date": "10/02/2025", "due": "12/03/2025", "num": "33446", "po": "", "dealer": "Kohler Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O212", "sku": "FBMG 24-26", "date": "03/13/2025", "due": "12/04/2025", "num": "33815", "po": "", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O213", "sku": "FBMG 24-26", "date": "04/21/2025", "due": "5/21/2025", "num": "34285", "po": "", "dealer": "S & S Motors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O214", "sku": "FBMG 24-26", "date": "4/28/2025", "due": "10/05/2025", "num": "34392", "po": "", "dealer": "JME Trailers, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O215", "sku": "FBMG 24-26", "date": "07/07/2025", "due": "06/08/2025", "num": "35130", "po": "", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O216", "sku": "FBMG 24-26", "date": "07/15/2025", "due": "8/14/2025", "num": "35211", "po": "", "dealer": "The Part Xperts", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O217", "sku": "FBMG 24-26", "date": "09/16/2025", "due": "10/31/2025", "num": "36022", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O218", "sku": "FBMG 24-26", "date": "09/16/2025", "due": "10/31/2025", "num": "36022", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O219", "sku": "FBMG 24-26", "date": "12/11/2025", "due": "12/12/2025", "num": "36767", "po": "3363", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O220", "sku": "FBMG 24-26", "date": "12/22/2025", "due": "01/21/2026", "num": "37184", "po": "", "dealer": "H5 Feed & Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O221", "sku": "FBMG 24-26", "date": "06/01/2026", "due": "05/02/2026", "num": "37279", "po": "Dean", "dealer": "Colt Bruegman Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O222", "sku": "FBMG 24-26", "date": "12/01/2026", "due": "11/02/2026", "num": "37361", "po": "3413", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O223", "sku": "FBMG 24-26", "date": "03/02/2026", "due": "05/03/2026", "num": "37655", "po": "", "dealer": "Kohler Trailer Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O224", "sku": "FBMG 24-26", "date": "09/02/2026", "due": "11/03/2026", "num": "37761", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O225", "sku": "FBMG 24-26", "date": "02/19/2026", "due": "3/21/2026", "num": "37877", "po": "", "dealer": "H5 Feed & Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O226", "sku": "FBMG 24-26", "date": "04/23/2026", "due": "05/23/2026", "num": "38719", "po": "3616", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O227", "sku": "FBMG 7.5-10", "date": "12/02/2026", "due": "03/14/2026", "num": "37802", "po": "", "dealer": "Loewen 1776 Outdoors LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O228", "sku": "FBMG 7.5-10", "date": "2/27/2026", "due": "3/29/2026", "num": "37987", "po": "", "dealer": "Riverside Boot & Saddle Blackfoot", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O229", "sku": "FBMG 7.5-10", "date": "4/13/2026", "due": "05/13/2026", "num": "38515", "po": "", "dealer": "Unlimited Equipment Sales, LLC", "qty": 3, "invoiced": 0, "backordered": 3}, {"id": "O230", "sku": "FBMG 99-02", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O231", "sku": "RBCGD 11-19", "date": "12/05/2026", "due": "11/06/2026", "num": "38946", "po": "", "dealer": "Open Range, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O232", "sku": "RBCGD 11-19", "date": "06/29/2026", "due": "07/29/2026", "num": "39464", "po": "3684", "dealer": "Truck Country", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O233", "sku": "RBCGD 20-23", "date": "6/22/2026", "due": "06/22/2026", "num": "39374", "po": "", "dealer": "King Gary", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O234", "sku": "RBCGD 20-23", "date": "06/23/2026", "due": "07/23/2026", "num": "39406", "po": "", "dealer": "Lone Star Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O235", "sku": "RBDD 10-23", "date": "02/25/2026", "due": "03/27/2026", "num": "37951", "po": "", "dealer": "Open Range, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O236", "sku": "RBDD 10-23", "date": "02/03/2026", "due": "01/04/2026", "num": "38006", "po": "", "dealer": "West Luke", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O237", "sku": "RBDD 10-23", "date": "03/30/2026", "due": "10/04/2026", "num": "38371", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O238", "sku": "RBDD 10-23", "date": "04/27/2026", "due": "05/27/2026", "num": "38766", "po": "Misenheimer", "dealer": "H5 Feed & Ranch, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O239", "sku": "RBDD 10-23", "date": "06/18/2026", "due": "07/18/2026", "num": "39346", "po": "", "dealer": "Scranton Truck & Trailer", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O240", "sku": "RBDD 10-23", "date": "06/23/2026", "due": "07/23/2026", "num": "39406", "po": "", "dealer": "Lone Star Trailers", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O241", "sku": "RBDD 10-23", "date": "01/07/2026", "due": "07/31/2026", "num": "39486", "po": "", "dealer": "Early Trailer Sales, LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O242", "sku": "RBDD 25-26", "date": "09/18/2024", "due": "10/18/2024", "num": "31985", "po": "", "dealer": "D & E Sales", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O243", "sku": "RBDD 25-26", "date": "12/13/2024", "due": "12/01/2025", "num": "32923", "po": "", "dealer": "D & E Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O244", "sku": "RBDD 25-26", "date": "09/16/2025", "due": "10/31/2025", "num": "36022", "po": "", "dealer": "C & J Traders", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O245", "sku": "RBDD 25-26", "date": "06/25/2026", "due": "10/07/2026", "num": "39428", "po": "", "dealer": "Walker Trailer and Equipment, LLC", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O246", "sku": "RBDD 95-02", "date": "12/11/2025", "due": "12/12/2025", "num": "36767", "po": "3363", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O247", "sku": "RBFD 08-16", "date": "02/25/2026", "due": "03/27/2026", "num": "37961", "po": "3512", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O248", "sku": "RBFD 08-16", "date": "08/06/2026", "due": "08/07/2026", "num": "39210", "po": "3664", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O249", "sku": "RBFD 17-22", "date": "12/01/2026", "due": "11/02/2026", "num": "37358", "po": "3412", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O250", "sku": "RBFD 17-22", "date": "01/28/2026", "due": "02/27/2026", "num": "37592", "po": "3466", "dealer": "Truck Country", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O251", "sku": "RBFD 23-26", "date": "08/15/2024", "due": "09/14/2024", "num": "31694", "po": "Colby Eli/Cody/Tyler/Dust", "dealer": "Route 66 Trailer Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O252", "sku": "RBFD 23-26", "date": "05/02/2025", "due": "07/03/2025", "num": "33386", "po": "", "dealer": "D & E Sales", "qty": 2, "invoiced": 0, "backordered": 2}, {"id": "O253", "sku": "RBFD 23-26", "date": "05/22/2025", "due": "10/06/2025", "num": "34692", "po": "", "dealer": "T & T Trailer Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}, {"id": "O254", "sku": "RBFD 23-26", "date": "02/13/2026", "due": "03/15/2026", "num": "37818", "po": "", "dealer": "Rod's Auto Sales LLC", "qty": 1, "invoiced": 0, "backordered": 1}]};

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
          <div className="sidebar-brand" style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '4px 12px 18px' }}>
            <div style={{ width: 30, height: 30, borderRadius: 6, background: '#2A3038', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
              <Truck size={16} color="#E8592A" />
            </div>
            <div>
              <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 16, letterSpacing: '0.01em', textTransform: 'uppercase', lineHeight: 1.1 }}>GR Bumpers</div>
              <div style={{ fontSize: 10.5, color: '#6B7280', marginTop: 2 }}>MX &amp; GR/US</div>
            </div>
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
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 20 }}>
          <div style={{ width: 32, height: 32, borderRadius: 7, background: '#1C2126', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Truck size={17} color="#E8592A" />
          </div>
          <div>
            <div style={{ fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: 18, textTransform: 'uppercase', lineHeight: 1.1 }}>Sign In</div>
            <div style={{ fontSize: 10.5, color: '#8A8F97' }}>Staff or dealer</div>
          </div>
        </div>
        <label style={{ fontSize: 11.5, fontWeight: 700, color: '#5B6470', textTransform: 'uppercase', letterSpacing: '0.03em', display: 'block', marginBottom: 5 }}>Username</label>
        <input value={username} onChange={e => setUsername(e.target.value)} placeholder="Leave blank if you're the owner" style={{
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
