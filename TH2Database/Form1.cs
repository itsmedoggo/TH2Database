using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json.Linq;

namespace TH2Database
{
    public partial class MainForm : Form
    {
        private ListViewColumnSorter monsterListSorter;
        private ListViewColumnSorter bossListSorter;
        private ListViewColumnSorter itemListSorter;
        private ListViewColumnSorter uniqueListSorter;
        private ListViewColumnSorter monsterBossesListSorter;
        private ListViewColumnSorter itemUniqueListSorter;
        private ListViewColumnSorter uniqueBossesListSorter;
        private ListViewColumnSorter setListSorter;
        private ListViewColumnSorter setUniquesListSorter;
        private ListViewColumnSorter affixListSorter;
        private ListViewColumnSorter rareAffixListSorter;
        public dynamic data;
        public string[] monsterTypes = { "Undead", "Demon", "Beast" };
        public Font plainFont = new Font("Courier New", 10f, FontStyle.Regular);
        public Font boldFont = new Font("Courier New", 10f, FontStyle.Bold);
        public Font italicFont = new Font("Courier New", 10f, FontStyle.Italic);
        public MainForm()
        {
            //string f = File.ReadAllText("data.json");
            //string f = File.ReadAllText("data.json");
            //data = JObject.Parse(f);
            var assembly = Assembly.GetExecutingAssembly();
            var file = "TH2Database.data.json";
            using (Stream s = assembly.GetManifestResourceStream(file))
            {
                using (StreamReader reader = new StreamReader(s))
                {
                    string f = reader.ReadToEnd();
                    data = JObject.Parse(f);
                }
            }

            InitializeComponent();

            monsterListSorter = new ListViewColumnSorter();
            bossListSorter = new ListViewColumnSorter();
            itemListSorter = new ListViewColumnSorter();
            uniqueListSorter = new ListViewColumnSorter();
            monsterBossesListSorter = new ListViewColumnSorter();
            itemUniqueListSorter = new ListViewColumnSorter();
            uniqueBossesListSorter = new ListViewColumnSorter();
            setListSorter = new ListViewColumnSorter();
            setUniquesListSorter = new ListViewColumnSorter();
            affixListSorter = new ListViewColumnSorter();
            rareAffixListSorter = new ListViewColumnSorter();
            monsterList.ListViewItemSorter = monsterListSorter;
            bossList.ListViewItemSorter = bossListSorter;
            itemList.ListViewItemSorter = itemListSorter;
            uniqueList.ListViewItemSorter = uniqueListSorter;
            monsterBossesList.ListViewItemSorter = monsterBossesListSorter;
            itemUniqueList.ListViewItemSorter = itemUniqueListSorter;
            uniqueBossesList.ListViewItemSorter = uniqueBossesListSorter;
            setList.ListViewItemSorter = setListSorter;
            setUniquesList.ListViewItemSorter = setUniquesListSorter;
            affixList.ListViewItemSorter = affixListSorter;
            rareAffixList.ListViewItemSorter = rareAffixListSorter;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            uniqueBossesClassBox.Items.Add("None");
            for (int i = 0; i < 29; i++)
            {
                uniqueBossesClassBox.Items.Add(getClassString(i));
            }
            uniqueBossesClassBox.SelectedIndex = 0;

            updateSelectionList(monsterList, data.monsters);
            updateSelectionList(bossList, data.bosses);
            updateSelectionList(itemList, data.items);
            updateSelectionList(uniqueList, data.uniques);
            updateSelectionList(setList, data.sets);
            updateSelectionList(affixList, data.affixes);
            updateSelectionList(rareAffixList, data.rareAffixes);
        }

        private void updateSelectionList(ListView list, JArray d)
        {
            list.Items.Clear();
            int id = 0;
            foreach (dynamic i in d)
            {
                ListViewItem item = new ListViewItem((string)i.name);
                item.SubItems.Add(id.ToString());
                list.Items.Add(item);
                id++;
            }
            list.Items[0].Selected = true;
        }

        private void sortColumn(ListView list, ListViewColumnSorter sorter, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == sorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (sorter.Order == SortOrder.Ascending)
                {
                    sorter.Order = SortOrder.Descending;
                }
                else
                {
                    sorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                sorter.SortColumn = e.Column;
                sorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            list.Sort();
        }

        private void monsterList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (monsterList.SelectedIndices.Count == 0) return;
            dynamic monster = data.monsters[monsterList.SelectedIndices[0]];
            monsterName.Text = monster.name;
            monsterType.Text = (string)(monsterTypes[(int)monster.type]);
            monsterLevelHorror.Text = monster.level.horror;
            monsterLevelPurgatory.Text = monster.level.purgatory;
            monsterLevelDoom.Text = monster.level.doom;
            monsterDungeonLvls.Text = monster.dungeonLvlMin == monster.dungeonLvlMax ? monster.dungeonLvlMin : monster.dungeonLvlMin + "..." + monster.dungeonLvlMax;
            monsterHealthHorror.Text = monster.healthMin.horror + "..." + monster.healthMax.horror;
            monsterHealthPurgatory.Text = monster.healthMin.purgatory + "..." + monster.healthMax.purgatory;
            monsterHealthDoom.Text = monster.healthMin.doom + "..." + monster.healthMax.doom;
            monsterHealthHorrorMP.Text = monster.healthMin.horrorMP + "..." + monster.healthMax.horrorMP;
            monsterHealthPurgatoryMP.Text = monster.healthMin.purgatoryMP + "..." + monster.healthMax.purgatoryMP;
            monsterHealthDoomMP.Text = monster.healthMin.doomMP + "..." + monster.healthMax.doomMP;
            monsterExpHorror.Text = monster.exp.horror;
            monsterExpPurgatory.Text = monster.exp.purgatory;
            monsterExpDoom.Text = monster.exp.doom;
            monsterAtk1DmgHorror.Text = monster.atk1DmgMin.horror + "..." + monster.atk1DmgMax.horror;
            monsterAtk1DmgPurgatory.Text = monster.atk1DmgMin.purgatory + "..." + monster.atk1DmgMax.purgatory;
            monsterAtk1DmgDoom.Text = monster.atk1DmgMin.doom + "..." + monster.atk1DmgMax.doom;
            monsterAtk1AccHorror.Text = monster.atk1Acc.horror;
            monsterAtk1AccPurgatory.Text = monster.atk1Acc.purgatory;
            monsterAtk1AccDoom.Text = monster.atk1Acc.doom;
            monsterAtk1Frames.Text = monster.atk1Frame;
            monsterSecondAtk.Text = monster.secondAtk == "true" ? "True" : "False";
            monsterAtk2DmgHorror.Text = monster.atk2DmgMin.horror + "..." + monster.atk2DmgMax.horror;
            monsterAtk2DmgPurgatory.Text = monster.atk2DmgMin.purgatory + "..." + monster.atk2DmgMax.purgatory;
            monsterAtk2DmgDoom.Text = monster.atk2DmgMin.doom + "..." + monster.atk2DmgMax.doom;
            monsterAtk2AccHorror.Text = monster.atk2Acc.horror;
            monsterAtk2AccPurgatory.Text = monster.atk2Acc.purgatory;
            monsterAtk2AccDoom.Text = monster.atk2Acc.doom;
            monsterAtk2Frames.Text = monster.atk2Frame;
            monsterArmorHorror.Text = monster.armor.horror;
            monsterArmorPurgatory.Text = monster.armor.purgatory;
            monsterArmorDoom.Text = monster.armor.doom;
            monsterAIType.Text = monster.aiType;
            monsterAISubtype.Text = monster.aiSubtype;
            monsterAIIQ.Text = monster.aiIQ;
            monsterSeedSize.Text = monster.seedSize;
            monsterDropType.Text = monster.dropType;
            monsterCanFly.Text = monster.canFly == "true" ? "True" : "False";
            setResistanceString((int)monster.resistances.horror[0], monsterRes11);
            setResistanceString((int)monster.resistances.horror[1], monsterRes12);
            setResistanceString((int)monster.resistances.horror[2], monsterRes13);
            setResistanceString((int)monster.resistances.horror[3], monsterRes14);
            setResistanceString((int)monster.resistances.horror[4], monsterRes15);
            setResistanceString((int)monster.resistances.horror[5], monsterRes16);
            setResistanceString((int)monster.resistances.doom[0], monsterRes21);
            setResistanceString((int)monster.resistances.doom[1], monsterRes22);
            setResistanceString((int)monster.resistances.doom[2], monsterRes23);
            setResistanceString((int)monster.resistances.doom[3], monsterRes24);
            setResistanceString((int)monster.resistances.doom[4], monsterRes25);
            setResistanceString((int)monster.resistances.doom[5], monsterRes26);
            populateMonsterBosses(monsterList.SelectedIndices[0]);
        }

        void setResistanceString(int resistance, Label label)
        {
            switch (resistance)
            {
                case 0:
                    label.Text = "Vulnerable";
                    label.Font = plainFont;
                    break;
                case 1:
                    label.Text = "Resistant";
                    label.Font = italicFont;
                    break;
                case 2:
                    label.Text = "Immune";
                    label.Font = boldFont;
                    break;
                default:
                    label.Text = resistance.ToString();
                    label.Font = boldFont;
                    break;
            }
        }

        private void bossList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (bossList.SelectedIndices.Count == 0) return;
            dynamic boss = data.bosses[Int32.Parse(bossList.SelectedItems[0].SubItems[1].Text)];
            bossName.Text = boss.name;
            bossBaseMonster.Text = (int)boss.packValue == 0 ? data.monsters[(int)boss.baseMonster].name : data.monsters[(int)boss.packValue].name;
            bossLevelHorror.Text = boss.level.horror;
            bossLevelPurgatory.Text = boss.level.purgatory;
            bossLevelDoom.Text = boss.level.doom;
            bossLevelHorrorMP.Text = boss.level.horrorMP;
            bossLevelPurgatoryMP.Text = boss.level.purgatoryMP;
            bossLevelDoomMP.Text = boss.level.doomMP;
            bossDungeonLevel.Text = (int)boss.dungeonLevel == 0 ? "Special" : boss.dungeonLevel;
            bossHealthHorror.Text = boss.health.horror;
            bossHealthPurgatory.Text = boss.health.purgatory;
            bossHealthDoom.Text = boss.health.doom;
            bossHealthHorrorMP.Text = boss.health.horrorMP;
            bossHealthPurgatoryMP.Text = boss.health.purgatoryMP;
            bossHealthDoomMP.Text = boss.health.doomMP;
            bossExpHorror.Text = boss.exp.horror;
            bossExpPurgatory.Text = boss.exp.purgatory;
            bossExpDoom.Text = boss.exp.doom;
            bossExpHorrorMP.Text = boss.exp.horrorMP;
            bossExpPurgatoryMP.Text = boss.exp.purgatoryMP;
            bossExpDoomMP.Text = boss.exp.doomMP;
            bossDmgHorror.Text = boss.dmgMin.horror + "..." + boss.dmgMax.horror;
            bossDmgPurgatory.Text = boss.dmgMin.purgatory + "..." + boss.dmgMax.purgatory;
            bossDmgDoom.Text = boss.dmgMin.doom + "..." + boss.dmgMax.doom;
            bossDmgHorrorMP.Text = boss.dmgMin.horrorMP + "..." + boss.dmgMax.horrormP;
            bossDmgPurgatoryMP.Text = boss.dmgMin.purgatoryMP + "..." + boss.dmgMax.purgatoryMP;
            bossDmgDoomMP.Text = boss.dmgMin.doomMP + "..." + boss.dmgMax.doomMP;
            bossAccHorror.Text = boss.acc.horror;
            bossAccPurgatory.Text = boss.acc.purgatory;
            bossAccDoom.Text = boss.acc.doom;
            bossAccHorrorMP.Text = boss.acc.horrorMP;
            bossAccPurgatoryMP.Text = boss.acc.purgatoryMP;
            bossAccDoomMP.Text = boss.acc.doomMP;
            bossArmorHorror.Text = boss.armor.horror;
            bossArmorPurgatory.Text = boss.armor.purgatory;
            bossArmorDoom.Text = boss.armor.doom;
            bossArmorHorrorMP.Text = boss.armor.horrorMP;
            bossArmorPurgatoryMP.Text = boss.armor.purgatoryMP;
            bossArmorDoomMP.Text = boss.armor.doomMP;
            setResistanceString((int)boss.resistances[0], bossRes1);
            setResistanceString((int)boss.resistances[1], bossRes2);
            setResistanceString((int)boss.resistances[2], bossRes3);
            setResistanceString((int)boss.resistances[3], bossRes4);
            setResistanceString((int)boss.resistances[4], bossRes5);
            setResistanceString((int)boss.resistances[5], bossRes6);
        }

        private void itemList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (itemList.SelectedIndices.Count == 0) return;
            dynamic item = data.items[Int32.Parse(itemList.SelectedItems[0].SubItems[1].Text)];
            itemName.Text = item.name;
            itemLevel.Text = item.level;
            itemDroppable.Text = item.droppable == 0 ? "False" : "True";
            setEquipSlotString((int)item.equipSlot, itemEquipSlot);
            setItemCategoryString((int)item.itemCategory, itemCategory);
            setItemTypeString((int)item.itemType, itemType);
            setSpecializationsString(item.specializations, itemSpecializations);
            itemDurabilityLabel.Text = item.hasQuantity == "false" ? "Durability:" : "Quantity:";
            itemPrice.Text = getValueRangeString(item.priceMin, item.priceMax);
            itemUnidentifiedName.Text = item.unidentifiedName == "" ? "Always identified" : item.unidentifiedName;
            itemDurability.Text = item.durMin == null ? "None" : getValueRangeString(item.durMin, item.durMax);
            itemArmor.Text = item.armorMin == null ? "None" : getValueRangeString(item.armorMin, item.armorMax);
            itemDFE.Text = item.dfeMin == null ? "None" : getValueRangeString(item.dfeMin, item.dfeMax);
            itemThornsDmg.Text = item.thornsDmgMinMin == null ? "None" : getValueRangeString(item.thornsDmgMinMin, item.thornsDmgMinMax, item.thornsDmgMaxMin, item.thornsDmgMaxMax);
            itemMeleeRes.Text = item.meleeResMinMin == null ? "None" : getValueRangeString(item.meleeResMinMin, item.meleeResMinMax, item.meleeResMaxMin, item.meleeResMaxMax, true);
            itemArrowRes.Text = item.arrowResMinMin == null ? "None" : getValueRangeString(item.arrowResMinMin, item.arrowResMinMax, item.arrowResMaxMin, item.arrowResMaxMax, true);
            itemDmg.Text = item.dmgMinMin == null ? "None" : getValueRangeString(item.dmgMinMin, item.dmgMinMax, item.dmgMaxMin, item.dmgMaxMax);
            itemStrReq.Text = item.strReqMin == null ? "0" : getValueRangeString(item.strReqMin, item.strReqMax);
            itemDexReq.Text = item.dexReqMin == null ? "0" : getValueRangeString(item.dexReqMin, item.dexReqMax);
            itemMagReq.Text = item.magReqMin == null ? "0" : getValueRangeString(item.magReqMin, item.magReqMax);
            itemVitReq.Text = item.vitReqMin == null ? "0" : getValueRangeString(item.vitReqMin, item.vitReqMax);
            itemLevelReq.Text = item.levelReqMin == null ? "0" : getValueRangeString(item.levelReqMin, item.levelReqMax);
            if (item.classReq == null) itemClassReq.Text = "None";
            else setItemClassReqString(item.classReq, itemClassReq);
            populateItemUniques((int)item.uID);
            if (item.additionalEffects == null) itemAdditionalEffects.Text = "None";
            else setAffixesString(item.additionalEffects, itemAdditionalEffects);
        }

        private void uniqueList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (uniqueList.SelectedIndices.Count == 0) return;
            dynamic unique = data.uniques[Int32.Parse(uniqueList.SelectedItems[0].SubItems[1].Text)];
            uniqueName.Text = unique.name;
            uniqueBaseItem.Text = "";
            int count = 0;
            foreach (dynamic item in data.items)
            {
                if (item.uID == unique.uID)
                {
                    if (count > 0 && count % 8 == 0) uniqueBaseItem.Text += "\n";
                    uniqueBaseItem.Text += item.name + ", ";
                    count++;
                }
            }
            if (uniqueBaseItem.Text == "") uniqueBaseItem.Text = "No valid base item (not droppable)";
            else uniqueBaseItem.Text = uniqueBaseItem.Text.Substring(0, uniqueBaseItem.Text.Length - 2);
            uniqueLevel.Text = unique.level;
            uniquePrice.Text = unique.price;
            populateUniqueBosses(unique);
            if (unique.additionalEffects == null) uniqueAdditionalEffects.Text = "None";
            else setAffixesString(unique.additionalEffects, uniqueAdditionalEffects);
            uniqueSet.Text = (int)unique.setID == -1 ? "None" : data.sets[(int)unique.setID].name;
        }

        private void setList_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (setList.SelectedIndices.Count == 0) return;
            dynamic set = data.sets[Int32.Parse(setList.SelectedItems[0].SubItems[1].Text)];
            setName.Text = set.name;
            setBonuses.Text = "";
            foreach (dynamic setBonus in set.setBonuses)
            {
                setBonuses.Text += "Required uniques: " + setBonus.requiredCount + "\n";
                setAffixesString(setBonus.effects, setBonuses, true, "  ");
                setBonuses.Text += "\n";
            }
            populateSetUniques(set);
        }

        private void affixList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (affixList.SelectedIndices.Count == 0) return;
            dynamic affix = data.affixes[Int32.Parse(affixList.SelectedItems[0].SubItems[1].Text)];
            affixName.Text = affix.name;
            affixSide.Text = (string)affix.prefix == "true" ? "Prefix" : "Suffix";
            affixLevel.Text = affix.level;
            affixPrice.Text = getValueRangeString(affix.priceMin, affix.priceMax);
            affixMultiplier.Text = affix.multiplier;
            affixRequiredLevel.Text = affix.requiredCharLevel;
            affixCursed.Text = affix.cursed;
            setSpecializationsString(affix.specializations, affixSpecializations);
            int count = 0;
            affixItemTypes.Text = "";
            foreach (int type in affix.itemTypes)
            {
                if (count > 0 && count % 8 == 0) affixItemTypes.Text += "\n";
                affixItemTypes.Text += getAffixItemTypeString(type);
                if (count < affix.itemTypes.Count) affixItemTypes.Text += ", ";
                count++;
            }
            setAffixesString(new JArray(affix.effects), affixEffects);
        }

        private void rareAffixList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rareAffixList.SelectedIndices.Count == 0) return;
            dynamic rareAffix = data.rareAffixes[Int32.Parse(rareAffixList.SelectedItems[0].SubItems[1].Text)];
            rareAffixName.Text = rareAffix.name;
            rareAffixSide.Text = (string)rareAffix.prefix == "true" ? "Prefix" : "Suffix";
            rareAffixLevel.Text = rareAffix.level;
            rareAffixPrice.Text = getValueRangeString(rareAffix.priceMin, rareAffix.priceMax);
            rareAffixMultiplier.Text = rareAffix.multiplier;
            rareAffixRequiredLevel.Text = rareAffix.requiredCharLevel;
            rareAffixCursed.Text = rareAffix.cursed;
            setSpecializationsString(rareAffix.specializations, rareAffixSpecializations);
            int count = 0;
            rareAffixItemTypes.Text = "";
            foreach (int type in rareAffix.itemTypes)
            {
                if (count > 0 && count % 8 == 0) rareAffixItemTypes.Text += "\n";
                rareAffixItemTypes.Text += getAffixItemTypeString(type);
                if (count < rareAffix.itemTypes.Count) rareAffixItemTypes.Text += ", ";
                count++;
            }
            setAffixesString(new JArray(rareAffix.effects), rareAffixEffects);
        }

        private void itemUniqueList_ItemActivate(object sender, EventArgs e)
        {
            uniqueList.Items[uniqueList.FindItemWithText(itemUniqueList.SelectedItems[0].Text).Index].Selected = true;
            uniqueList.Items[uniqueList.SelectedIndices[0]].EnsureVisible();
            tabs.SelectedTab = uniquesTab;
        }

        private void monsterBossesList_ItemActivate(object sender, EventArgs e)
        {
            bossList.Items[bossList.FindItemWithText(monsterBossesList.SelectedItems[0].Text).Index].Selected = true;
            bossList.Items[bossList.SelectedIndices[0]].EnsureVisible();
            tabs.SelectedTab = bossesTab;
        }

        private void uniqueBossesList_ItemActivate(object sender, EventArgs e)
        {
            bossList.Items[bossList.FindItemWithText(uniqueBossesList.SelectedItems[0].Text).Index].Selected = true;
            bossList.Items[bossList.SelectedIndices[0]].EnsureVisible();
            tabs.SelectedTab = bossesTab;
        }

        private void setUniquesList_ItemActivate(object sender, EventArgs e)
        {
            uniqueList.Items[uniqueList.FindItemWithText(setUniquesList.SelectedItems[0].Text).Index].Selected = true;
            uniqueList.Items[uniqueList.SelectedIndices[0]].EnsureVisible();
            tabs.SelectedTab = uniquesTab;
        }

        string getValueRangeString(dynamic valMin, dynamic valMax, bool percent = false)
        {
            return (string)(valMin == valMax ? valMin : valMin + (percent ? "%" : "") + "..." + valMax) + (percent ? "%" : "");
        }

        string getValueRangeString(int valMinI, int valMaxI, bool percent = false)
        {
            string valMin = valMinI.ToString();
            string valMax = valMaxI.ToString();
            return (valMin == valMax ? valMin : valMin + (percent ? "%" : "") + "..." + valMax) + (percent ? "%" : "");
        }

        string getValueRangeString(double valMinI, double valMaxI, bool percent = false)
        {
            string valMin = valMinI.ToString();
            string valMax = valMaxI.ToString();
            return (valMin == valMax ? valMin : valMin + (percent ? "%" : "") + "..." + valMax) + (percent ? "%" : "");
        }

        string getValueRangeString(dynamic val1Min, dynamic val1Max, dynamic val2Min, dynamic val2Max, bool percent = false)
        {
            /*if (val1Min == val1Max && val2Min == val2Max)
            {
                return getValueRangeString(val1Min, val2Min, percent);
            }*/
            return getValueRangeString(val1Min, val1Max, percent) + " - " + getValueRangeString(val2Min, val2Max, percent);
        }

        void setEquipSlotString(int slot, Label label)
        {
            switch (slot)
            {
                case 0:
                    label.Text = "One-Handed";
                    break;
                case 1:
                    label.Text = "Two-Handed";
                    break;
                case 2:
                    label.Text = "Body";
                    break;
                case 3:
                    label.Text = "Head";
                    break;
                case 4:
                    label.Text = "Ring";
                    break;
                case 5:
                    label.Text = "Amulet";
                    break;
                case 6:
                    label.Text = "Unequippable";
                    break;
                case 8:
                    label.Text = "Belt";
                    break;
                case 9:
                    label.Text = "Gloves";
                    break;
                case 10:
                    label.Text = "Boots";
                    break;
                default:
                    label.Text = slot.ToString();
                    break;
            }
        }

        void setItemCategoryString(int category, Label label)
        {
            switch (category)
            {
                case 1:
                    label.Text = "Weapon";
                    break;
                case 2:
                    label.Text = "Armor";
                    break;
                case 3:
                    label.Text = "Jewelry/Potion/Relict";
                    break;
                case 4:
                    label.Text = "Gold";
                    break;
                case 5:
                    label.Text = "Quest Item";
                    break;
                default:
                    label.Text = category.ToString();
                    break;
            }
        }

        void setItemTypeString(int type, Label label, bool append = false)
        {
            if (!append) label.Text = "";
            switch (type)
            {
                case 0:
                    label.Text += "Misc.";
                    break;
                case 1:
                    label.Text += "Sword";
                    break;
                case 2:
                    label.Text += "Axe";
                    break;
                case 3:
                    label.Text += "Bow";
                    break;
                case 4:
                    label.Text += "Blunt";
                    break;
                case 5:
                    label.Text += "Shield/Mage Offhand";
                    break;
                case 6:
                    label.Text += "Light Armor";
                    break;
                case 7:
                    label.Text += "Helm";
                    break;
                case 8:
                    label.Text += "Medium Armor";
                    break;
                case 9:
                    label.Text += "Heavy Armor";
                    break;
                case 10:
                    label.Text += "Staff";
                    break;
                case 11:
                    label.Text += "Gold";
                    break;
                case 12:
                    label.Text += "Ring";
                    break;
                case 13:
                    label.Text += "Amulet";
                    break;
                case 15:
                    label.Text += "Gloves";
                    break;
                case 16:
                    label.Text += "Boots";
                    break;
                case 17:
                    label.Text += "Belt";
                    break;
                case 18:
                    label.Text += "Flask";
                    break;
                case 19:
                    label.Text += "Trap";
                    break;
                case 20:
                    label.Text += "Claw";
                    break;
                case 21:
                    label.Text += "Knife";
                    break;
                case 22:
                    label.Text += "Throwing Mallet";
                    break;
                default:
                    label.Text += type.ToString();
                    break;
            }
        }

        string getAffixItemTypeString(int type)
        {
            switch (type)
            {
                case 0:
                    return "One-Handed Sword";
                case 1:
                    return "Two-Handed Sword";
                case 2:
                    return "Axe";
                case 3:
                    return "One-Handed Blunt";
                case 4:
                    return "Two-Handed Blunt";
                case 5:
                    return "Bow";
                case 6:
                    return "Flask";
                case 7:
                    return "Shield";
                case 8:
                    return "Light Armor";
                case 9:
                    return "Medium Armor";
                case 10:
                    return "Heavy Armor";
                case 11:
                    return "Helm";
                case 12:
                    return "Belt";
                case 13:
                    return "Gloves";
                case 15:
                    return "Boots";
                case 16:
                    return "Staff";
                case 17:
                    return "Ring";
                case 18:
                    return "Amulet";
                case 19:
                    return "Claw";
                case 20:
                    return "Trap";
                default:
                    return type.ToString();
            }
        }

        void setSpecializationsString(JArray specs, Label label)
        {
            label.Text = "";
            foreach (int spec in specs)
            {
                switch (spec)
                {
                    case 0:
                        label.Text += "Battle\n";
                        break;
                    case 1:
                        label.Text += "Magic\n";
                        break;
                    case 2:
                        label.Text += "Thorns\n";
                        break;
                    case 3:
                        label.Text += "Summoning\n";
                        break;
                    default:
                        label.Text += spec.ToString() + "\n";
                        break;
                }
            }
        }

        void setItemClassReqString(JArray reqs, Label label)
        {
            label.Text = "";
            if (reqs.Count == 29)
            {
                label.Text = "Any";
                return;
            }
            int count = 0;
            foreach (int req in reqs)
            {
                if (count > 0 && count % 8 == 0)
                {
                    label.Text += "\n";
                }
                label.Text += getClassString(req);
                count++;
                if (count < reqs.Count)
                {
                    label.Text += ", ";
                }
            }
        }

        string getClassString(int i)
        {
            switch (i)
            {
                case 0:
                    return "Warrior";
                case 1:
                    return "Inquisitor";
                case 2:
                    return "Guardian";
                case 3:
                    return "Templar";
                case 4:
                    return "Archer";
                case 5:
                    return "Scout";
                case 6:
                    return "Sharpshooter";
                case 7:
                    return "Trapper";
                case 8:
                    return "Mage";
                case 9:
                    return "Elementalist";
                case 10:
                    return "Demonologist";
                case 11:
                    return "Necromancer";
                case 12:
                    return "Beastmaster";
                case 13:
                    return "Warlock";
                case 14:
                    return "Monk";
                case 15:
                    return "Kensei";
                case 16:
                    return "Shugoki";
                case 17:
                    return "Shinobi";
                case 18:
                    return "Rogue";
                case 19:
                    return "Assassin";
                case 20:
                    return "Iron Maiden";
                case 21:
                    return "Bombardier";
                case 22:
                    return "Savage";
                case 23:
                    return "Berserker";
                case 24:
                    return "Executioner";
                case 25:
                    return "Thraex";
                case 26:
                    return "Murmillo";
                case 27:
                    return "Dimachaerus";
                case 28:
                    return "Secutor";
                default:
                    return i.ToString();
            }
        }

        void populateItemUniques(int uID)
        {
            itemUniqueList.Items.Clear();
            if (uID == 0) return;
            foreach (dynamic unique in data.uniques)
            {
                if ((int)unique.uID == uID)
                {
                    ListViewItem item = new ListViewItem((string)unique.name);
                    item.SubItems.Add((string)unique.level);
                    itemUniqueList.Items.Add(item);
                }
            }
        }

        void populateMonsterBosses(int index)
        {
            monsterBossesList.Items.Clear();
            foreach (dynamic boss in data.bosses)
            {
                int baseMonsterID = (int)boss.packValue == 0 ? (int)boss.baseMonster : (int)boss.packValue;
                if (baseMonsterID == index)
                {
                    ListViewItem item = new ListViewItem((string)boss.name);
                    item.SubItems.Add((string)boss.dungeonLevel);
                    monsterBossesList.Items.Add(item);
                }
            }
        }

        void populateUniqueBosses(dynamic unique)
        {
            uniqueBossesList.Items.Clear();
            dynamic item = null;
            int selectedClass = uniqueBossesClassBox.SelectedIndex - 1;
            foreach (dynamic i in data.items)
            {
                if (i.uID == unique.uID)
                {
                    if (i.classReq == null)
                    {
                        item = i;
                        break;
                    }
                    foreach (dynamic c in i.classReq)
                    {
                        if ((int)c == selectedClass || selectedClass == -1)
                        {
                            item = i;
                            break;
                        }
                    }
                }
            }
            if (item == null || (int)item.droppable == 0) return;
            Dictionary<int, string> horrorDict = new Dictionary<int, string>();
            Dictionary<int, string> purgatoryDict = new Dictionary<int, string>();
            Dictionary<int, string> doomDict = new Dictionary<int, string>();
            foreach (dynamic boss in data.bosses)
            {
                int baseMonsterID = (int)boss.packValue == 0 ? (int)boss.baseMonster : (int)boss.packValue;
                int baseMonsterLevel = (int)data.monsters[baseMonsterID].level.horror;
                int bossLevelHorror = baseMonsterLevel;
                int bossLevelPurgatory = baseMonsterLevel;
                int bossLevelDoom = baseMonsterLevel;
                if (boss.notes != null)
                {
                    switch ((string)boss.name)
                    {
                        case "Skeleton King":
                            bossLevelHorror = 24;
                            break;
                        case "The Butcher":
                            bossLevelHorror = 16;
                            break;
                        case "Lich King":
                            bossLevelHorror = 60;
                            bossLevelPurgatory = 60;
                            bossLevelDoom = 60;
                            break;
                        case "Mordessa":
                            bossLevelHorror = 59;
                            bossLevelPurgatory = 59;
                            bossLevelDoom = 59;
                            break;
                        case "Wielder of Shadowfang":
                            bossLevelHorror = 60;
                            bossLevelPurgatory = 60;
                            bossLevelDoom = 60;
                            break;
                        case "Hephasto the Armorer":
                            bossLevelHorror = 61;
                            bossLevelPurgatory = 61;
                            bossLevelDoom = 61;
                            break;
                    }
                }
                //Horror
                if ((int)unique.level <= bossLevelHorror + 4 && itemWithinQLvlRange((int)item.level, bossLevelHorror, 0) && underStatReqLimit(item, (int)boss.dungeonLevel, 0) && !isHigherLevelValidUnique(unique, bossLevelHorror))
                {
                    ListViewItem i = new ListViewItem((string)boss.name);
                    i.SubItems.Add("Horror");
                    i.SubItems.Add((int)boss.dungeonLevel == 0 ? "Special" : (string)boss.dungeonLevel);
                    if (!horrorDict.ContainsKey(bossLevelHorror)) horrorDict.Add(bossLevelHorror, getDropOdds(item, bossLevelHorror, 0, (int)boss.dungeonLevel));
                    i.SubItems.Add(horrorDict[bossLevelHorror]);
                    uniqueBossesList.Items.Add(i);
                }
                //Purgatory
                if ((int)unique.level <= bossLevelPurgatory + 4 && itemWithinQLvlRange((int)item.level, bossLevelPurgatory, 1) && underStatReqLimit(item, (int)boss.dungeonLevel, 1) && validItemLevelForDifficulty(item, 1) && !isHigherLevelValidUnique(unique, bossLevelPurgatory))
                {
                    ListViewItem i = new ListViewItem((string)boss.name);
                    i.SubItems.Add("Purgatory");
                    i.SubItems.Add((int)boss.dungeonLevel == 0 ? "Special" : (string)boss.dungeonLevel);
                    if (!purgatoryDict.ContainsKey(bossLevelPurgatory)) purgatoryDict.Add(bossLevelPurgatory, getDropOdds(item, bossLevelPurgatory, 1, (int)boss.dungeonLevel));
                    i.SubItems.Add(purgatoryDict[bossLevelPurgatory]);
                    uniqueBossesList.Items.Add(i);
                }
                //Doom
                if ((int)unique.level <= bossLevelDoom + 4 && itemWithinQLvlRange((int)item.level, bossLevelDoom, 2) && validItemLevelForDifficulty(item, 2) && !isHigherLevelValidUnique(unique, bossLevelDoom))
                {
                    ListViewItem i = new ListViewItem((string)boss.name);
                    i.SubItems.Add("Doom");
                    i.SubItems.Add((int)boss.dungeonLevel == 0 ? "Special" : (string)boss.dungeonLevel);
                    if (!doomDict.ContainsKey(bossLevelDoom)) doomDict.Add(bossLevelDoom, getDropOdds(item, bossLevelDoom, 2, (int)boss.dungeonLevel));
                    i.SubItems.Add(doomDict[bossLevelDoom]);
                    uniqueBossesList.Items.Add(i);
                }
            }
        }

        void populateSetUniques(dynamic set)
        {
            setUniquesList.Items.Clear();
            foreach (dynamic uID in set.uIDs)
            {
                dynamic unique = data.uniques[(int)uID];
                ListViewItem item = new ListViewItem((string)unique.name);
                item.SubItems.Add((string)unique.level);
                setUniquesList.Items.Add(item);
            }
        }

        bool itemWithinQLvlRange(int level, int monsterLevel, int difficulty)
        {
            /*int minLvl = 3 + (int)((monsterLevel + 3) / 6) + 10 * difficulty;
            int maxLvl = 6 + (int)((monsterLevel + 3) / 3) + 20 * difficulty;
            return level <= maxLvl && level >= minLvl;*/
            return true;
        }

        bool underStatReqLimit(dynamic baseItem, int dungeonLevel, int difficulty)
        {
            int limit;
            if (difficulty == 0) limit = dungeonLevel * 5 + 20;
            else if (difficulty == 1) limit = dungeonLevel * 5 + 135;
            else return true;
            if (baseItem.strReqMax != null && (int)baseItem.strReqMax > limit) return false;
            if (baseItem.dexReqMax != null && (int)baseItem.dexReqMax > limit) return false;
            if (baseItem.magReqMax != null && (int)baseItem.magReqMax > limit) return false;
            if (baseItem.vitReqMax != null && (int)baseItem.vitReqMax > limit) return false;
            return true;
        }

        string getDropOdds(dynamic baseItem, int bossLevel, int difficulty, int dlvl)
        {
            if (uniqueBossesClassBox.SelectedIndex == 0) return " ";
            int validDrops = 0;
            int selectedClass = uniqueBossesClassBox.SelectedIndex - 1;
            bool correctBaseItemClass = false;
            bool correctItemClass = false;
            if (baseItem.classReq == null) correctBaseItemClass = true;
            else
            {
                foreach (dynamic c in baseItem.classReq)
                {
                    if ((int)c == selectedClass) correctBaseItemClass = true;
                }
            }
            if (!correctBaseItemClass) return " ";
            foreach (dynamic u in data.uniques)
            {
                dynamic item = null;
                correctItemClass = false;
                foreach (dynamic i in data.items)
                {
                    if (i.uID == u.uID)
                    {
                        if (i.classReq == null)
                        {
                            correctItemClass = true;
                            item = i;
                            break;
                        }
                        foreach (dynamic c in i.classReq)
                        {
                            if ((int)c == selectedClass)
                            {
                                correctItemClass = true;
                                item = i;
                                break;
                            }
                        }
                    }
                }
                if (correctItemClass && (int)item.droppable != 0 && (int)u.level <= bossLevel + 4 && (int)item.level <= bossLevel && validItemLevelForDifficulty(item, difficulty) && underStatReqLimit(item, dlvl, difficulty) && !isHigherLevelValidUnique(u, bossLevel)) validDrops++;
            }
            return (100d / validDrops).ToString("0.00") + "%";
        }

        bool isHigherLevelValidUnique(dynamic unique, int bossLevel)
        {
            foreach (dynamic u in data.uniques)
            {
                if (u.uID == unique.uID && (int)u.level > (int)unique.level && (int)u.level <= bossLevel + 4) return true;
            }
            return false;
        }

        bool validItemLevelForDifficulty(dynamic item, int difficulty)
        {
            if (difficulty == 0) return true;
            int t = (int)item.itemType;
            if (difficulty == 1)
            {
                switch (t)
                {
                    case 1: //Sword
                        return item.level >= 7;
                    case 2: //Axe
                        return item.level >= 7;
                    case 3: //Bow
                        return item.level >= 11;
                    case 4: //Mace
                        return item.level >= 12;
                    case 5: //Shield
                        return item.level >= 6;
                    case 6: //Light armor
                        return item.level >= 8;
                    case 7: //Helm
                        return item.level >= 8;
                    case 10: //Staff
                        return item.level >= 7;
                    case 15: //Glove
                        return item.level >= 9;
                    case 16: //Boots
                        return item.level >= 9;
                    case 17: //Belt
                        return item.level >= 7;
                    case 18: //Flask
                        return item.level >= 12;
                    case 19: //Trap
                        return item.level >= 12;
                    case 20: //Claw
                        return item.level >= 11;
                    default:
                        return true;
                }
            }
            else
            {
                switch (t)
                {
                    case 1: //Sword
                        return item.level >= 30;
                    case 2: //Axe
                        return item.level >= 25;
                    case 3: //Bow
                        return item.level >= 20;
                    case 4: //Mace
                        return item.level >= 30;
                    case 5: //Shield
                        return item.level >= 12;
                    case 6: //Light armor
                        return false;
                    case 7: //Helm
                        return item.level >= 12;
                    case 10: //Staff
                        return item.level >= 19;
                    case 15: //Glove
                        return item.level >= 28;
                    case 16: //Boots
                        return item.level >= 23;
                    case 17: //Belt
                        return item.level >= 26;
                    case 18: //Flask
                        return item.level >= 24;
                    case 19: //Trap
                        return item.level >= 24;
                    case 20: //Claw
                        return item.level >= 25;
                    default:
                        return true;
                }
            }
        }

        void setAffixesString(JArray affixes, Label label, bool append = false, string prefix = "")
        {
            if (!append) label.Text = "";
            foreach (dynamic affix in affixes)
            {
                label.Text += prefix;
                switch ((int)affix.effectID)
                {
                    case 0:
                        label.Text += "Accuracy: " + getValueRangeString(affix.effectMin1, affix.effectMax1);
                        break;
                    case 1:
                        label.Text += "Damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 2:
                        label.Text += "Damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        label.Text += ", Accuracy: " + getValueRangeString(affix.effectMin1, affix.effectMax1);
                        break;
                    case 3:
                        label.Text += "Armor: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 4:
                        label.Text += "Armor set to: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 5:
                        label.Text += "Armor: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 6:
                        label.Text += "Fire resistance: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 7:
                        label.Text += "Lightning resistance: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 8:
                        label.Text += "Magic resistance: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 9:
                        label.Text += "Resist all: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 11:
                        label.Text += "Spell levels: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 12:
                        label.Text += "Extra charges: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 13:
                        label.Text += getSpellString((int)affix.effectMin2) + "charges: " + affix.effectMax2;
                        break;
                    case 14:
                        label.Text += "Fire melee damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 15:
                        label.Text += "Lightning melee damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 16:
                        label.Text += "Strength: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 17:
                        label.Text += "Magic: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 18:
                        label.Text += "Dexterity: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 19:
                        label.Text += "Vitality: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 20:
                        label.Text += "Attributes: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 21:
                        label.Text += "Damage from enemies: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 22:
                        label.Text += "Hit points: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 23:
                        label.Text += "Mana: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 24:
                        label.Text += "Durability: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 25:
                        label.Text += "Durability: " + getValueRangeString(-(int)affix.effectMin2, -(int)affix.effectMax2, true);
                        break;
                    case 27:
                        label.Text += "Light radius: " + getValueRangeString((int)affix.effectMin2 * 10, (int)affix.effectMax2 * 10, true);
                        break;
                    case 28:
                        label.Text += "Light radius: " + getValueRangeString(-(int)affix.effectMin2 * 10, -(int)affix.effectMax2 * 10, true);
                        break;
                    case 29:
                        label.Text += "Multishot";
                        break;
                    case 30:
                        label.Text += "Fire arrow damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 31:
                        label.Text += "Lightning arrow damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 32:
                        label.Text += "Fireball damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 33:
                        label.Text += "Damage against undead: 40%";
                        break;
                    case 36:
                        label.Text += "Arrow resistance: 33%";
                        break;
                    case 37:
                        label.Text += "Knocks target back";
                        break;
                    case 38:
                        label.Text += "Damage against demons: 40%";
                        break;
                    case 39:
                        label.Text += "Lose all resistances";
                        break;
                    case 40:
                        label.Text += "Increased fury/concentration duration";
                        break;
                    case 41:
                        if (affix.effectMin2 == 3) label.Text += "Melee hits leech 1% mana";
                        else label.Text += "Melee hits leech 2% mana";
                        break;
                    case 42:
                        if (affix.effectMin2 == 3) label.Text += "Melee hits leech 1% life";
                        else label.Text += "Melee hits leech 2% life";
                        break;
                    case 43:
                        label.Text += "Armor piercing: " + getValueRangeString((double)affix.effectMin2 * 12.5, (double)affix.effectMax2 * 12.5, true);
                        break;
                    case 44:
                        switch ((int)affix.effectMin2)
                        {
                            case 1:
                                label.Text += "Quick";
                                break;
                            case 2:
                                label.Text += "Fast";
                                break;
                            case 3:
                                label.Text += "Faster";
                                break;
                            case 4:
                                label.Text += "Fastest";
                                break;
                        }
                        label.Text += " attack speed";
                        break;
                    case 45:
                        switch ((int)affix.effectMin2)
                        {
                            case 1:
                                label.Text += "Fast";
                                break;
                            case 2:
                                label.Text += "Faster";
                                break;
                            case 3:
                                label.Text += "Fastest";
                                break;
                        }
                        label.Text += " hit recovery";
                        break;
                    case 46:
                        label.Text += "Fast block";
                        break;
                    case 47:
                        label.Text += "Damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 48:
                        label.Text += "Randomized arrow speed";
                        break;
                    case 49:
                        label.Text += "Damage set to: " + getValueRangeString(affix.effectMin2, affix.effectMin2, affix.effectMax2, affix.effectMax2);
                        break;
                    case 50:
                        label.Text += "Durability set to: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 53:
                        label.Text += "Life regeneration: 100%";
                        break;
                    case 54:
                        label.Text += "Increased life stealing";
                        break;
                    case 55:
                        label.Text += "No strength requirement";
                        break;
                    case 56:
                        label.Text += "Enhanced perception";
                        break;
                    case 57:
                        label.Text += "Unique graphic";
                        break;
                    case 58:
                        label.Text += "Lightning damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 60:
                        label.Text += "30% chance of dealing +60% damage";
                        break;
                    case 62:
                        label.Text += "Mana regeneration: 100%";
                        break;
                    case 63:
                        label.Text += "Randomized damage: 80% - 160%";
                        break;
                    case 64:
                        label.Text += "Damage: " + getValueRangeString(140 + (int)affix.effectMin2 * 2, 140 + (int)affix.effectMax2 * 2, true);
                        label.Text += ", Durability: " + getValueRangeString(-(int)affix.effectMin2, -(int)affix.effectMax2, true);
                        break;
                    case 65:
                        label.Text += "Mana cost: -25%";
                        break;
                    case 66:
                        label.Text += "Armor against demons: 20";
                        break;
                    case 67:
                        label.Text += "Armor against undead: 15";
                        break;
                    case 68:
                        label.Text += "50% mana moved to health";
                        break;
                    case 69:
                        label.Text += "40% health moved to mana";
                        break;
                    case 70:
                        label.Text += "Gold find: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 71:
                        label.Text += "Gold find: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 72:
                        label.Text += "Magic find: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 73:
                        if (affix.element == null) label.Text += "Error: bugged affix is missing spell"; //Agonizer bugged affix
                        else label.Text += affix.element == 0 ? getSpellString(1) : getSpellString((int)System.Math.Log(System.Math.Abs((int)affix.element), 2)) + " spell level: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 74:
                        label.Text += affix.element == 0 ? getSpellString(1) : getSpellString((int)System.Math.Log(System.Math.Abs((int)affix.element), 2)) + " damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 75:
                        label.Text += (int)affix.element == 0 ? getSpellString(1) : getSpellString((int)System.Math.Log(System.Math.Abs((long)affix.element), 2)) + " damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 76:
                        label.Text += getElementString((int)affix.element) + " damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 77:
                        label.Text += (affix.effectMin1 == null ? "100%" : getValueRangeString(affix.effectMin1, affix.effectMax1, true)) + " chance to add " + getElementString((int)affix.element) + " damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 78:
                        label.Text += getValueRangeString(affix.effectMin1, affix.effectMax1, true) + " chance to add " + getElementString((int)affix.element) + " damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 79:
                        label.Text += getElementString((int)affix.element) + " damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 80:
                        label.Text += getElementString((int)affix.element) + " damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 81:
                        label.Text += "Magic arrow damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 82:
                        label.Text += "Holy arrow damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 83:
                        label.Text += "Acid arrow damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 84:
                        switch ((int)affix.effectMin2)
                        {
                            case 1:
                                label.Text += "Fast";
                                break;
                            case 2:
                                label.Text += "Faster";
                                break;
                            case 3:
                                label.Text += "Fastest";
                                break;
                        }
                        label.Text += " cast speed";
                        break;
                    case 85:
                        label.Text += "All attributes: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 86:
                        label.Text += "Strength: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 87:
                        label.Text += "Dexterity: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 88:
                        label.Text += "Magic: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 89:
                        label.Text += "Vitality: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 90:
                        label.Text += "Life regeneration: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 91:
                        label.Text += "Mana regeneration: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 92:
                        label.Text += "Life regeneration: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 93:
                        label.Text += "Mana regeneration: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 94:
                        label.Text += "Experience: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 95:
                        label.Text += "Experience: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 96:
                        label.Text += "Melee resistance: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 97:
                        label.Text += "Arrow resistance: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 98:
                        label.Text += "Minion health: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 99:
                        label.Text += "Minion damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 100:
                        label.Text += "Minion armor: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 101:
                        label.Text += "Minion accuracy: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 102:
                        label.Text += "Minion health: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 103:
                        label.Text += "Minion damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 104:
                        label.Text += "Minion armor: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 105:
                        label.Text += "Block chance: " + getValueRangeString(affix.effectMin1, affix.effectMax1, true);
                        break;
                    case 106:
                        label.Text += "Crit chance: " + getValueRangeString(affix.effectMin1, affix.effectMax1, true);
                        break;
                    case 107:
                        label.Text += "Crit damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 108:
                        label.Text += "Crit chance: " + getValueRangeString(affix.effectMin1, affix.effectMax1, true);
                        label.Text += ", Crit damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 109:
                        label.Text += "Crit damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 110:
                        label.Text += "Damage against " + getMonsterTypeString((int)affix.element) + ": " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 111:
                        if (affix.element == null) label.Text += "Error: bugged affix is missing enemy type"; //Blistered heart of beastrage bugged affix (fixed in upcoming version)
                        else label.Text += "Damage against " + getMonsterTypeString((int)affix.element) + ": " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 112:
                        label.Text += "Accuracy against " + getMonsterTypeString((int)affix.element) + ": " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 113:
                        label.Text += "Armor against " + getMonsterTypeString((int)affix.element) + ": " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 114:
                        label.Text += "Armor against " + getMonsterTypeString((int)affix.element) + ": " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 115:
                        label.Text += "Damage from " + getMonsterTypeString((int)affix.element) + ": " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 116:
                        label.Text += "Health: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 117:
                        label.Text += "Mana: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 119:
                        label.Text += getElementString((int)System.Math.Pow(2, (int)affix.element)) + " resistance: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 122:
                        label.Text += "Thorns damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, true);
                        break;
                    case 123:
                        label.Text += "Stun threshold: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 124:
                        label.Text += "Arcane melee damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 125:
                        label.Text += "Holy melee damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 126:
                        label.Text += "Acid melee damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 127:
                        label.Text += "Level requirement: " + getValueRangeString(affix.effectMin2, affix.effectMax2);
                        break;
                    case 128:
                        label.Text += "Cold melee damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    case 129:
                        label.Text += "Cold arrow damage: " + getValueRangeString(affix.effectMin2, affix.effectMax2, affix.effectMaxMin2, affix.effectMaxMax2);
                        break;
                    default: 
                        label.Text += "Unknown effect " + (string)affix.effectID;
                        break;
                }
                label.Text += "\n";
            }
        }

        string getSpellString(int id)
        {
            switch (id)
            {
                case 1:
                    return "Fire Bolt";
                case 2:
                    return "Healing";
                case 3:
                    return "Lightning";
                case 4:
                    return "Flash";
                case 5:
                    return "Identify";
                case 6:
                    return "Fire Wall";
                case 7:
                    return "Town Portal";
                case 8:
                    return "Stone Curse";
                case 9:
                    return "Seeing";
                case 10:
                    return "Phasing";
                case 11:
                    return "Mana Shield";
                case 12:
                    return "Fire Blast";
                case 13:
                    return "Hydra";
                case 14:
                    return "Ball Lightning";
                case 15:
                    return "Force Wave";
                case 16:
                    return "Reflect";
                case 17:
                    return "Lightning Ring";
                case 18:
                    return "Lightning Nova";
                case 19:
                    return "Flame Ring";
                case 20:
                    return "Inferno";
                case 21:
                    return "Golem";
                case 22:
                    return "Fury";
                case 23:
                    return "Teleport";
                case 24:
                    return "Apocalypse";
                case 25:
                    return "Ethereal";
                case 26:
                    return "Item Repair";
                case 27:
                    return "Staff Recharge";
                case 28:
                    return "Trap Disarm";
                case 29:
                    return "Elemental";
                case 30:
                    return "Charged Bolt";
                case 31:
                    return "Holy Bolt";
                case 33:
                    return "Telekinesis";
                case 34:
                    return "Heal Other";
                case 35:
                    return "Arcane Star";
                case 36:
                    return "Bone Spirit";
                case 37:
                    return "Mana";
                case 38:
                    return "the Magi?";
                case 39:
                    return "Holy Nova";
                case 40:
                    return "Lightning Wall";
                case 41:
                    return "Fire Nova";
                case 42:
                    return "Warp";
                case 43:
                    return "Arcane Nova";
                case 44:
                    return "Berserk";
                case 45:
                    return "Ring of Fire";
                case 52:
                    return "Tier 1 Summon";
                case 53:
                    return "Tier 2 Summon";
                case 54:
                    return "Tier 3 Summon";
                case 55:
                    return "Unsummon";
                default:
                    return "unk_" + id.ToString();
            }
        }

        string getElementString(int id)
        {
            switch (id)
            {
                case 1:
                    return "Physical";
                case 2:
                    return "Fire";
                case 3:
                    return "Fire Arrow";
                case 4:
                    return "Lightning";
                case 8:
                    return "Magic";
                case 9:
                    return "Thorns";
                case 16:
                    return "Acid";
                case 32:
                    return "Holy";
                case 64:
                    return "Cold";
                default:
                    return "unk_" + id.ToString();
            }
        }

        string getMonsterTypeString(int id)
        {
            switch (id)
            {
                case 1:
                    return "undead";
                case 2:
                    return "demons";
                case 4:
                    return "beasts";
                default:
                    return "unkmonstertype_" + id.ToString();
            }
        }

        private void monsterList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(monsterList, monsterListSorter, e);
        }

        private void bossList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(bossList, bossListSorter, e);
        }

        private void itemList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(itemList, itemListSorter, e);
        }

        private void uniqueList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(uniqueList, uniqueListSorter, e);
        }

        private void monsterBossesList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(monsterBossesList, monsterBossesListSorter, e);
        }

        private void itemUniqueList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(itemUniqueList, itemUniqueListSorter, e);
        }

        private void uniqueBossesList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(uniqueBossesList, uniqueBossesListSorter, e);
        }

        private void setList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(setList, setListSorter, e);
        }

        private void setUniquesList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(setUniquesList, setUniquesListSorter, e);
        }

        private void affixList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(affixList, affixListSorter, e);
        }

        private void rareAffixList_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            sortColumn(rareAffixList, rareAffixListSorter, e);
        }

        private void uniqueBossesClassBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (uniqueList.SelectedIndices.Count == 0) return;
            populateUniqueBosses(data.uniques[Int32.Parse(uniqueList.SelectedItems[0].SubItems[1].Text)]);
        }

    }
}
