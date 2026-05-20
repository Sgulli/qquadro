export const mockClubs = [
  {
    id: "club-1",
    name: "Rosso Neri",
    seats: [
      { originalClubId: "club-1", sectorId: "sector-a1" },
      { originalClubId: "club-1", sectorId: "sector-a1" },
    ],
  },
  {
    id: "club-2",
    name: "Club Azzurro",
    seats: [{ originalClubId: "club-2", sector: { macroAreaId: "ma-curva" } }],
  },
  {
    id: "club-3",
    name: "Rosa Nera",
    seats: [
      { originalClubId: "club-3", sectorId: "sector-b1" },
      { originalClubId: "club-3", sectorId: "sector-b1" },
      { originalClubId: "club-3", sectorId: "sector-b1" },
    ],
  },
  { id: "club-4", name: "Viola Club", seats: [] },
  {
    id: "club-5",
    name: "GialloBlu Ultras",
    seats: [{ originalClubId: "club-5", sectorId: "sector-b1" }],
  },
];

export const mockMacroareas = [
  {
    id: "ma-tribuna",
    name: "Tribuna Centrale",
    sectors: [
      { id: "sector-a1", name: "Settore A" },
      { id: "sector-a2", name: "Settore B" },
    ],
    saleInclusions: [{ macroAreaId: "ma-tribuna", saleId: "sale-1" }],
  },
  {
    id: "ma-curva",
    name: "Curva Nord",
    sectors: [],
    saleInclusions: [{ macroAreaId: "ma-curva", saleId: "sale-1" }],
  },
  {
    id: "ma-distinti",
    name: "Distinti",
    sectors: [{ id: "sector-b1", name: "Settore C" }],
    saleInclusions: [{ macroAreaId: "ma-distinti", saleId: "sale-1" }],
  },
];

export const mockSeatRequests = [
  {
    clubId: "club-1",
    sectorId: "sector-a1",
    sector: { name: "Settore A" },
    quantity: 12,
    macroAreaId: "ma-tribuna",
    macroArea: { name: "Tribuna Centrale" },
  },
  {
    clubId: "club-1",
    sectorId: "sector-a2",
    sector: { name: "Settore B" },
    quantity: 8,
    macroAreaId: "ma-tribuna",
    macroArea: { name: "Tribuna Centrale" },
  },
  {
    clubId: "club-1",
    sectorId: null,
    sector: null,
    quantity: 5,
    macroAreaId: "ma-curva",
    macroArea: { name: "Curva Nord" },
  },
  {
    clubId: "club-2",
    sectorId: "sector-a1",
    sector: { name: "Settore A" },
    quantity: 15,
    macroAreaId: "ma-tribuna",
    macroArea: { name: "Tribuna Centrale" },
  },
  {
    clubId: "club-2",
    sectorId: "sector-a2",
    sector: { name: "Settore B" },
    quantity: 6,
    macroAreaId: "ma-tribuna",
    macroArea: { name: "Tribuna Centrale" },
  },
  {
    clubId: "club-2",
    sectorId: null,
    sector: null,
    quantity: 10,
    macroAreaId: "ma-curva",
    macroArea: { name: "Curva Nord" },
  },
  {
    clubId: "club-3",
    sectorId: "sector-a1",
    sector: { name: "Settore A" },
    quantity: 20,
    macroAreaId: "ma-tribuna",
    macroArea: { name: "Tribuna Centrale" },
  },
  {
    clubId: "club-3",
    sectorId: "sector-b1",
    sector: { name: "Settore C" },
    quantity: 14,
    macroAreaId: "ma-distinti",
    macroArea: { name: "Distinti" },
  },
  {
    clubId: "club-3",
    sectorId: null,
    sector: null,
    quantity: 7,
    macroAreaId: "ma-curva",
    macroArea: { name: "Curva Nord" },
  },
  {
    clubId: "club-4",
    sectorId: "sector-a2",
    sector: { name: "Settore B" },
    quantity: 10,
    macroAreaId: "ma-tribuna",
    macroArea: { name: "Tribuna Centrale" },
  },
  {
    clubId: "club-4",
    sectorId: "sector-b1",
    sector: { name: "Settore C" },
    quantity: 6,
    macroAreaId: "ma-distinti",
    macroArea: { name: "Distinti" },
  },
  {
    clubId: "club-5",
    sectorId: "sector-a1",
    sector: { name: "Settore A" },
    quantity: 9,
    macroAreaId: "ma-tribuna",
    macroArea: { name: "Tribuna Centrale" },
  },
  {
    clubId: "club-5",
    sectorId: "sector-b1",
    sector: { name: "Settore C" },
    quantity: 11,
    macroAreaId: "ma-distinti",
    macroArea: { name: "Distinti" },
  },
  {
    clubId: "club-5",
    sectorId: null,
    sector: null,
    quantity: 3,
    macroAreaId: "ma-curva",
    macroArea: { name: "Curva Nord" },
  },
];

export const mockResultFreeSeats = [
  { freeSeats: 80 },
  { freeSeats: 60 },
  { freeSeats: 40 },
  { freeSeats: 50 },
];
