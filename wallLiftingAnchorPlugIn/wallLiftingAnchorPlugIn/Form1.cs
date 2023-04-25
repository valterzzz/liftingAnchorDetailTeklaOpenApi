using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Model;

namespace wallLiftingAnchorPlugIn
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            comboBox1.SelectedItem = formWorkType;
            textBox1.Enabled = false;
            textBox1.Text = calculateAdhesionForce(comboBox1.Text);

            comboBox2.SelectedItem = typeOfElement;
            textBox2.Enabled = false;
            textBox2.Text = calculateAdhesionFactor(comboBox2.Text);

            comboBox3.SelectedItem = typeOfTransportation;
            textBox3.Enabled = false;
            textBox3.Text = calculateDynamicFactor(comboBox3.Text);

            comboBox4.SelectedItem = cableAngle;
            textBox4.Enabled = false;
            textBox4.Text = calculateAngleFactor(comboBox4.Text);

            comboBox5.SelectedItem = numberOfAnchors;

        }

        //Create parameters
        private static string formWorkType = "Oiled steel formwork, oiled plastic-coated plywood";

        private static string typeOfElement = "Standard";

        private static string typeOfTransportation = "Tower crane, portal crane and mobile crane";

        private static string cableAngle = "30.0";

        private static string numberOfAnchors = "2";


        class TerwaSATTU
        {
            public List<Hashtable> terwaSATTU = new List<Hashtable>();

            public TerwaSATTU()
            {
                //Create a hashable for every lifting anchor and add to Array /ALWAYS START WITH WEAKEST ANCHOR/

                Hashtable terwaSATTU14 = new Hashtable();

                terwaSATTU14.Add("Anchor Type", "SA-TTU-014-200-black");
                terwaSATTU14.Add("Axial pull 100% <30degrees", "14");
                terwaSATTU14.Add("Diognal pull 80% 30<degrees<45", "11");
                terwaSATTU14.Add("Diognal pull and Angled pull 100% <45", "14");
                terwaSATTU14.Add("Tilting 50%", "7");

                terwaSATTU.Add(terwaSATTU14);

                Hashtable terwaSATTU25 = new Hashtable();

                terwaSATTU25.Add("Anchor Type", "SA-TTU-025-230-black");
                terwaSATTU25.Add("Axial pull 100% <30degrees", "25");
                terwaSATTU25.Add("Diognal pull 80% 30<degrees<45", "20");
                terwaSATTU25.Add("Diognal pull and Angled pull 100% <45", "25");
                terwaSATTU25.Add("Tilting 50%", "13");

                terwaSATTU.Add(terwaSATTU25);

                Hashtable terwaSATTU40 = new Hashtable();

                terwaSATTU40.Add("Anchor Type", "SA-TTU-040-270-black");
                terwaSATTU40.Add("Axial pull 100% <30degrees", "38");
                terwaSATTU40.Add("Diognal pull 80% 30<degrees<45", "30");
                terwaSATTU40.Add("Diognal pull and Angled pull 100% <45", "40");
                terwaSATTU40.Add("Tilting 50%", "20");

                terwaSATTU.Add(terwaSATTU40);

                Hashtable terwaSATTU50 = new Hashtable();

                terwaSATTU50.Add("Anchor Type", "SA-TTU-050-290-black");
                terwaSATTU50.Add("Axial pull 100% <30degrees", "47");
                terwaSATTU50.Add("Diognal pull 80% 30<degrees<45", "38");
                terwaSATTU50.Add("Diognal pull and Angled pull 100% <45", "50");
                terwaSATTU50.Add("Tilting 50%", "25");

                terwaSATTU.Add(terwaSATTU50);

                Hashtable terwaSATTU75 = new Hashtable();

                terwaSATTU75.Add("Anchor Type", "SA-TTU-075-320-black");
                terwaSATTU75.Add("Axial pull 100% <30degrees", "65");
                terwaSATTU75.Add("Diognal pull 80% 30<degrees<45", "52");
                terwaSATTU75.Add("Diognal pull and Angled pull 100% <45", "75");
                terwaSATTU75.Add("Tilting 50%", "38");

                terwaSATTU.Add(terwaSATTU75);

                Hashtable terwaSATTU100 = new Hashtable();

                terwaSATTU100.Add("Anchor Type", "SA-TTU-100-390-black");
                terwaSATTU100.Add("Axial pull 100% <30degrees", "85");
                terwaSATTU100.Add("Diognal pull 80% 30<degrees<45", "68");
                terwaSATTU100.Add("Diognal pull and Angled pull 100% <45", "100");
                terwaSATTU100.Add("Tilting 50%", "50");

                terwaSATTU.Add(terwaSATTU100);

                Hashtable terwaSATTU125 = new Hashtable();

                terwaSATTU125.Add("Anchor Type", "SA-TTU-125-500-black");
                terwaSATTU125.Add("Axial pull 100% <30degrees", "120");
                terwaSATTU125.Add("Diognal pull 80% 30<degrees<45", "96");
                terwaSATTU125.Add("Diognal pull and Angled pull 100% <45", "125");
                terwaSATTU125.Add("Tilting 50%", "62.5");

                terwaSATTU.Add(terwaSATTU125);

                Hashtable terwaSATTU170 = new Hashtable();

                terwaSATTU170.Add("Anchor Type", "SA-TTU-170-500-Black");
                terwaSATTU170.Add("Axial pull 100% <30degrees", "140");
                terwaSATTU170.Add("Diognal pull 80% 30<degrees<45", "110");
                terwaSATTU170.Add("Diognal pull and Angled pull 100% <45", "170");
                terwaSATTU170.Add("Tilting 50%", "85");

                terwaSATTU.Add(terwaSATTU170);

                Hashtable terwaSATTU220 = new Hashtable();

                terwaSATTU220.Add("Anchor Type", "SA-TTU-220-500-Black");
                terwaSATTU220.Add("Axial pull 100% <30degrees", "200");
                terwaSATTU220.Add("Diognal pull 80% 30<degrees<45", "160");
                terwaSATTU220.Add("Diognal pull and Angled pull 100% <45", "220");
                terwaSATTU220.Add("Tilting 50%", "110");

                terwaSATTU.Add(terwaSATTU220);

            }

            public int CalculatePullDetail(double concreteCubeStrenght, double forceOnAnchor, double angleOfPull)
            {
                int indexOfSelectedAnchor = -1;

                //Calculate lifting anchor for Axial/Angled Pull

                if (concreteCubeStrenght >= 25.00)
                {
                    foreach (var obj in terwaSATTU)
                    {
                        if (double.Parse(obj["Diognal pull and Angled pull 100% <45"].ToString()) >= forceOnAnchor)
                        {
                            indexOfSelectedAnchor = terwaSATTU.IndexOf(obj);

                            return indexOfSelectedAnchor;
                        }
                    }
                }
                else if (angleOfPull <= 30.00 & concreteCubeStrenght >= 15.00)
                {
                    foreach (var obj in terwaSATTU)
                    {
                        if (double.Parse(obj["Axial pull 100% <30degrees"].ToString()) >= forceOnAnchor)
                        {
                            indexOfSelectedAnchor = terwaSATTU.IndexOf(obj);

                            return indexOfSelectedAnchor;
                        }
                    }
                }
                else if (angleOfPull >= 30.00 & angleOfPull <= 45.00 & concreteCubeStrenght >= 15.00)
                {
                    foreach (var obj in terwaSATTU)
                    {
                        if (double.Parse(obj["Diognal pull 80% 30<degrees<45"].ToString()) >= forceOnAnchor)
                        {
                            indexOfSelectedAnchor = terwaSATTU.IndexOf(obj);

                            return indexOfSelectedAnchor;
                        }
                    }
                }

                return indexOfSelectedAnchor;

            }
            public int CalculateTiltUpDetail(double concreteCubeStrenght, double forceOnAnchor)
            {
                int indexOfSelectedAnchor = -1;

                if (concreteCubeStrenght >= 15.00)
                {
                    foreach (var obj in terwaSATTU)
                    {
                        if (double.Parse(obj["Tilting 50%"].ToString()) >= forceOnAnchor)
                        {
                            indexOfSelectedAnchor = terwaSATTU.IndexOf(obj);

                            return indexOfSelectedAnchor;
                        }
                    }
                }

                return indexOfSelectedAnchor;
            }  



        }



        //Function to calculate surface adhesion foce
        private static string calculateAdhesionForce(string typeOfSurface)
        {
            string adhesionForce;
            if (typeOfSurface == "Oiled steel formwork, oiled plastic-coated plywood")
            {
                adhesionForce = "1";
            }
            else if (typeOfSurface == "Varnished timber formwork with panel boards")
            {
                adhesionForce = "2";
            }
            else
            {
                adhesionForce = "3";
            }

            return adhesionForce;
        }

        //Function to calculate adhesion multiplier
        private static string calculateAdhesionFactor(string typeOfElement)
        {
            string adhesionFactor;
            if (typeOfElement == "Standard")
            {
                adhesionFactor = "1";
            }
            else if (typeOfElement == "𝝅 - panels")
            {
                adhesionFactor = "2";
            }
            else if (typeOfElement == "Ribbed elements")
            {
                adhesionFactor = "3";
            }
            else
            {
                adhesionFactor = "4";
            }

            return adhesionFactor;
        }

        //Function to calculate dynamic factor
        private static string calculateDynamicFactor(string typeOfTransportation)
        {
            string dynamicFactor;
            if (typeOfTransportation == "Tower crane, portal crane and mobile crane")
            {
                dynamicFactor = "1.3";
            }
            else if (typeOfTransportation == "Lifting and moving on flat terrain")
            {
                dynamicFactor = "2.5";
            }
            else
            {
                dynamicFactor = "4";
            }

            return dynamicFactor;
        }

        //Function to calculate cable angle factor
        private static string calculateAngleFactor(string cableAngle)
        {

            double angle = double.Parse(cableAngle);

            double angleFactor = 1 / Math.Cos(DegreesToRadians(angle));

            angleFactor = Math.Round(angleFactor, 3);

            return angleFactor.ToString();

        }

        //Function to convert Degrees into Radians
        private static double DegreesToRadians(double degrees)
        {
            return degrees * Math.PI / 180.0;
        }


        //Function to check if the model is connected 
        private static bool CheckIfModelIsConnected(Tekla.Structures.Model.Model model)
        {
            bool isConnected = true;

            if (!model.GetConnectionStatus())
            {
                MessageBox.Show("Tekla Structures not connected!");
                isConnected = false;
            }
            return isConnected;
        }

        //Function to check if array has only Part (Beam/PolyBeam/ContourPlate)

        private static bool CheckIfOnlyPart(ModelObjectEnumerator elementArray)
        {
            bool isPart = true;
            foreach (Tekla.Structures.Model.Object obj in elementArray)
            {
                if (obj.GetType().Name != "Beam" & obj.GetType().Name != "PolyBeam" & obj.GetType().Name != "ContourPlate")
                {
                    MessageBox.Show("Non-Parts Selected!");
                    MessageBox.Show(obj.GetType().Name);
                    isPart = false;
                }
                
            }

            return isPart;
        }

        //Function to calculate De-mold/Tilt-up at the plant force

        private static double CalculateForceFactoryTiltUp(double weightOfElement, double adhesionForce, double adhesionFactor, double areaOfElement, double angleFactor, double numberOfAnchors)
        {

            double forceInTiltUp = ((weightOfElement / 2 + adhesionForce * adhesionFactor * areaOfElement) * angleFactor) / numberOfAnchors;

            return forceInTiltUp;
        }

        //Function to calculate Transport at the plant force

        private static double CalculateForceTransport(double weightOfElement, double dynamicFactor, double angleFactor, double numberOfAnchors)
        {
            double forceInFactoryTransport = (weightOfElement * dynamicFactor * angleFactor) / numberOfAnchors;

            return forceInFactoryTransport;
        }



        private void insertButton_Click(object sender, EventArgs e)
        {
            Model model = new Model();

            if (CheckIfModelIsConnected(model) == false)
            {
                return;
            }
            

            //Get selected objects
            Tekla.Structures.Model.UI.ModelObjectSelector mS = new Tekla.Structures.Model.UI.ModelObjectSelector();

            Tekla.Structures.Model.ModelObjectEnumerator selectedObjectList = mS.GetSelectedObjects();

            if (CheckIfOnlyPart(selectedObjectList) == false)
            {
                return;
            }


            //Additional List of Elements Because the First One is Not usable
            Tekla.Structures.Model.ModelObjectEnumerator selectedObjectList2 = mS.GetSelectedObjects();

            foreach (Tekla.Structures.Model.Part objPart in selectedObjectList2)
            {

                Tekla.Structures.Model.Assembly obj = objPart.GetAssembly();
                //Get Calculation properties

                //Get element weight
                double elementWeight = 0.00;
                obj.GetReportProperty("WEIGHT", ref elementWeight);

                elementWeight = elementWeight * 0.01;

                //Get element length

                double elementLength = 0.00;
                obj.GetReportProperty("LENGTH", ref elementLength);

                //Get element area

                double elementArea = 0.00;
                obj.GetReportProperty("AREA_PROJECTION_XY_NET", ref elementArea);

                elementArea = elementArea * 1e-6;

                //Get concrete class from element main part

                string materialString = string.Empty;
                obj.GetReportProperty("MAINPART.MATERIAL", ref materialString);

                string[] splitted = materialString.Split('/');

                //Assume at least 30/37 concrete at 70% capacity

                double material = 25.90;
                
                if (splitted.Length >= 2)
                {
                    material = double.Parse(splitted[splitted.Length - 1]);
                }
                
                //Get adhesion force

                double adhesionForce = double.Parse(textBox1.Text);

                //Get adhesion shape factor

                double adhesionFactor = double.Parse(textBox2.Text);

                //Get dynamic factor 

                double dynamicFactor = double.Parse(textBox3.Text);

                //Get cable angle factor

                double cableAngleFactor = double.Parse(textBox4.Text);

                //Get lifting anchor count 

                double liftingAnchorCount = double.Parse(comboBox5.Text);

                //Calculate required anchor 
                TerwaSATTU liftingAnchor = new TerwaSATTU();


                //Calculate the force in anchor from De-mold/Tilt-up at the plant

                double tiltUpForcekN = Math.Round(CalculateForceFactoryTiltUp(elementWeight, adhesionForce, adhesionFactor, elementArea, cableAngleFactor, liftingAnchorCount), 2);

                int keyFactoryTiltUp = liftingAnchor.CalculateTiltUpDetail(material, tiltUpForcekN);

                //Calculate the force in anchor from transportation at the plant and calculate the anchor

                double transportForceFactorykN = Math.Round(CalculateForceTransport(elementWeight, dynamicFactor, cableAngleFactor, liftingAnchorCount), 2);

                int keyTransportFactoryPull = liftingAnchor.CalculatePullDetail(material, transportForceFactorykN, double.Parse(cableAngle));


                //Calculate the force in anchor from transportation at the site

                double transportForceSitekN = Math.Round(CalculateForceTransport(elementWeight, dynamicFactor, cableAngleFactor, liftingAnchorCount), 2);

                int keyTransportSitePull = liftingAnchor.CalculatePullDetail(material, transportForceSitekN, double.Parse(cableAngle));


                //Find the required lifting anchor

                List<int> requiredLiftingAnchorList = new List<int>();

                requiredLiftingAnchorList.Add(keyFactoryTiltUp);
                requiredLiftingAnchorList.Add(keyTransportFactoryPull);
                requiredLiftingAnchorList.Add(keyTransportSitePull);

                int isLiftable = requiredLiftingAnchorList.Min();


                //Check if each selection gave back at least possible lifting anchor

                string liftingAnchorString = string.Empty;

                if (isLiftable != -1)
                {
                    int requiredLiftingAnchor = requiredLiftingAnchorList.Max();

                    Hashtable calculatedLiftingAnchor = new Hashtable();

                    calculatedLiftingAnchor = liftingAnchor.terwaSATTU[requiredLiftingAnchor];

                    liftingAnchorString = calculatedLiftingAnchor["Anchor Type"].ToString();

                    //MessageBox.Show(calculatedLiftingAnchor["Anchor Type"].ToString());
                }
                else
                {
                    MessageBox.Show("Key not Selectable For This Element");
                    continue;
                }

                //Add lifting anchor inside the Assembly

                Tekla.Structures.Model.Component liftingAnchorComponent = new Tekla.Structures.Model.Component();

                liftingAnchorComponent.Name = "Lifting anchor";
                liftingAnchorComponent.Number = 30000080;

                Tekla.Structures.Model.ComponentInput CI = new ComponentInput();

                CI.AddInputObject(objPart);

                liftingAnchorComponent.SetComponentInput(CI);
                liftingAnchorComponent.LoadAttributesFromFile("standard");

                //Change attributes of the custom component to reflect design
                liftingAnchorComponent.SetAttribute("custom", "1");
                liftingAnchorComponent.SetAttribute("DistOutOfBeam", "0");
                liftingAnchorComponent.SetAttribute("custom_profile", liftingAnchorString);

                //Add in lifting anchors based on count

                if (liftingAnchorCount == 2.00)
                {
                    if (!liftingAnchorComponent.Insert())
                    {
                        Console.WriteLine("Component Insert Failed!");
                    }
                    else
                    {
                        Console.WriteLine(liftingAnchorComponent.Identifier.ID);
                    }
                }
                else if (liftingAnchorCount == 4.00)
                {
                    //Use procentage
                    liftingAnchorComponent.SetAttribute("distanceUnit", "1");
                    liftingAnchorComponent.SetAttribute("distCOGEndPerc", "20.00");
                    liftingAnchorComponent.SetAttribute("distCOGStartPerc", "20.00");

                    if (!liftingAnchorComponent.Insert())
                    {
                        Console.WriteLine("Component Insert Failed!");
                    }
                    else
                    {
                        Console.WriteLine(liftingAnchorComponent.Identifier.ID);
                    }

                    liftingAnchorComponent.SetAttribute("distCOGEndPerc", "40.00");
                    liftingAnchorComponent.SetAttribute("distCOGStartPerc", "40.00");

                    if (!liftingAnchorComponent.Insert())
                    {
                        Console.WriteLine("Component Insert Failed!");
                    }
                    else
                    {
                        Console.WriteLine(liftingAnchorComponent.Identifier.ID);
                    }

                }



                //Comit changes to model

                model.CommitChanges();


            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {
            numberOfAnchors = comboBox5.Text;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            cableAngle = comboBox4.Text;
            textBox4.Text = calculateAngleFactor(comboBox4.Text);
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            formWorkType = comboBox1.Text;
            textBox1.Text = calculateAdhesionForce(comboBox1.Text);
        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            typeOfElement = comboBox2.Text;
            textBox2.Text = calculateAdhesionFactor(comboBox2.Text);
        }

        private void comboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            typeOfTransportation = comboBox3.Text;
            textBox3.Text = calculateDynamicFactor(comboBox3.Text);
        }

        private void comboBox5_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

    }
}
