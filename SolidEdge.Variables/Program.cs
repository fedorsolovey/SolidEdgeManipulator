// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.

// Copyright (c) Siemens Product Lifecycle Management Software Inc. All rights reserved.

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using SolidEdgeFramework;

namespace SolidEdge.Part.Variables
{
    public class VariableObject
    {
        public string name { get; set; }
        public double value { get; set; }
    }

    class Program
    {
        static VariableList variableList;

        [STAThread]
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeFramework.Documents documents = null;
            SolidEdgeFramework.SolidEdgeDocument document = null;
            SolidEdgeFramework.Variables variables = null;
            
            try
            {
                Console.WriteLine("Registering OleMessageFilter.");

                // Register with OLE to handle concurrency issues on the current thread.
                OleMessageFilter.Register();

                Console.WriteLine("Connecting to Solid Edge.");

                // Connect to or start Solid Edge.
                application = ConnectToSolidEdge(true);

                // Make sure user can see the GUI.
                application.Visible = true;

                // Bring Solid Edge to the foreground.
                application.Activate();

                // Get a reference to the Documents collection.
                documents = application.Documents;

                // This check is necessary because application.ActiveDocument will throw an
                // exception if no documents are open...
                if (documents.Count > 0)
                {
                    // Attempt to connect to ActiveDocument.
                    document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
                }

                // Make sure we have a document.
                if (document == null)
                {
                    throw new System.Exception("No active document.");
                }

                variables = (SolidEdgeFramework.Variables)document.Variables;
                variableList = VariableList(variables);

                writeToFile(@"\\Mac\Home\Downloads\Windows\output.txt", variableList);

                var variableObject = readFile(@"\\Mac\Home\Downloads\Windows\input.txt");

            }
            catch (System.Exception ex)
            {
#if DEBUG
                System.Diagnostics.Debugger.Break();
#endif
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Unregistering OleMessageFilter.");
                OleMessageFilter.Revoke();

                //Console.ReadLine();
            }
        }

        static VariableList VariableList(SolidEdgeFramework.Variables variables)
        {
            return (SolidEdgeFramework.VariableList)variables.Query(
                pFindCriterium: "*",
                NamedBy: SolidEdgeConstants.VariableNameBy.seVariableNameByBoth,
                VarType: SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth);
        }

        static VariableList ProcessVariables(SolidEdgeFramework.Variables variables) 
        {
            SolidEdgeFramework.VariableList variableList = null;
            //SolidEdgeFramework.variable variable = null;
            //SolidEdgeFrameworkSupport.Dimension dimension = null;
            //dynamic variableListItem = null;    // In C#, the dynamic keyword is used so we don't have to call InvokeMember().

            // Get a reference to the variablelist.
            variableList = (SolidEdgeFramework.VariableList)variables.Query(
                pFindCriterium: "*",
                NamedBy: SolidEdgeConstants.VariableNameBy.seVariableNameByBoth,
                VarType: SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth);

            return variableList;

            // Process variables.
            //for (int i = 1; i <= variableList.Count; i++)
            //{
            //    // Get a reference to variable item.
            //    variableListItem = variableList.Item(i);

            //    // Determine the variable item type.
            //    SolidEdgeConstants.ObjectType objectType = (SolidEdgeConstants.ObjectType)variableListItem.Type;

            //    Console.WriteLine("{0} = {1}", variableListItem.VariableTableName, variableListItem.Value);

            //    // Process the specific variable item type.
            //    switch (objectType)
            //    {
            //        case SolidEdgeConstants.ObjectType.igDimension:
            //            dimension = (SolidEdgeFrameworkSupport.Dimension)variableListItem;
            //            break;
            //        case SolidEdgeConstants.ObjectType.igVariable:
            //            variable = (SolidEdgeFramework.variable)variableListItem;
            //            break;
            //        default:
            //            // Other SolidEdgeConstants.ObjectType's may exist.
            //            break;
            //    }
            //}
        }

        public static void writeToFile(string path, VariableList variableList)
        {
            var tableVariables = new List<VariableObject>();
            dynamic variableListItem = null;

            for (int i = 1; i <= variableList.Count; i++)
            {
                variableListItem = variableList.Item(i);

                var variableObject = new VariableObject()
                {
                    name = variableListItem.VariableTableName,
                    value = variableListItem.Value
                };
                tableVariables.Add(variableObject);
            }

            string output = JsonConvert.SerializeObject(tableVariables);
            System.IO.File.WriteAllText(path, output);
        }

        public static VariableObject readFile(string path)
        {
            string input = System.IO.File.ReadAllText(path);
            var variable = JsonConvert.DeserializeObject<VariableObject>(input);
            return variable;
        }

        /// <summary>
        /// Connects to a running instance of Solid Edge.
        /// </summary>
        public static SolidEdgeFramework.Application ConnectToSolidEdge()
        {
            return ConnectToSolidEdge(false);
        }

        /// <summary>
        /// Connects to a running instance of Solid Edge with an option to start if not running.
        /// </summary>
        public static SolidEdgeFramework.Application ConnectToSolidEdge(bool startIfNotRunning)
        {
            try
            {
                // Attempt to connect to a running instance of Solid Edge.
                return (SolidEdgeFramework.Application)
                    Marshal.GetActiveObject("SolidEdge.Application");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // Failed to connect.
                if (ex.ErrorCode == -2147221021 /* MK_E_UNAVAILABLE */)
                {
                    if (startIfNotRunning)
                    {
                        // Start Solid Edge.
                        return (SolidEdgeFramework.Application)
                            Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
                    }
                    else
                    {
                        throw new System.Exception("Solid Edge is not running.");
                    }
                }
                else
                {
                    throw;
                }
            }
            catch
            {
                throw;
            }
        }    
    }
}


