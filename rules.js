export const MATERIAL_RULES = [
  {active:true, materialA:'MS', materialB:'MILD STEEL', reason:'General'},
  {active:true, materialA:'Perspex', materialB:'BLACK PERSPEX', reason:'General'},
  {active:true, materialA:'MS', materialB:'Mild Steel Sheet', reason:'General'},
  {active:true, materialA:'Armox 500', materialB:'Armox', reason:'General'},
  {active:true, materialA:'Armox 500', materialB:'Ramor 500', thicknessRule:4, reason:'Approved 4 mm substitution'},
  {active:true, materialA:'ALTread', materialB:'Al Tread 1200H4', reason:'General'},
  {active:true, materialA:'Strenx 700', materialB:'Strenx', reason:'General'},
  {active:true, materialA:'Chromodeck', materialB:'Chromodeck Sheet', reason:'General'},
  {active:true, materialA:'SS', materialB:'Gr304 SS', reason:'General'},
  {active:true, materialA:'Ramor 500', materialB:'Ramort', reason:'General'},
  {active:true, materialA:'Ramor 500', materialB:'Ramort 500', reason:'General'},
  {active:true, materialA:'ALTread', materialB:'Aluminium Tread', reason:'General'},
  {active:true, materialA:'Hardox', materialB:'Hardox 500', reason:'General'},
  {active:true, materialA:'Ramor 500', materialB:'Ramor® 500', reason:'General'},
  {active:true, materialA:'MS', materialB:'Mild Steel Plate', reason:'General'},
  {active:true, materialA:'Aluminium', materialB:'Aluminium Sheet', reason:'General'},
  {active:true, materialA:'ALTread', materialB:'AL Chequer Plate', reason:'General'},
  {active:true, materialA:'Gr304 SS', materialB:'GR 304 SS', reason:'General'}
];
export const THICKNESS_RULES = [
  {active:true, materialFilter:'Armox', fromThickness:8.5, toThickness:9, reason:'Spec 10/31: 8.5mm Armox treated as 9mm'},
  {active:true, materialFilter:'Ramor', fromThickness:8.5, toThickness:9, reason:'Spec 10/31: 8.5mm Ramor treated as 9mm'},
  {active:true, materialFilter:'', fromThickness:3.5, toThickness:4, reason:'General half-mm normalization'}
];
export function cloneDefaultRules(){return {materialRules:MATERIAL_RULES.map(x=>({...x})), thicknessRules:THICKNESS_RULES.map(x=>({...x}))};}
