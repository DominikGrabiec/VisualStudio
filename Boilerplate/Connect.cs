﻿using System;
using System.IO;
using System.Resources;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;

namespace Boilerplate
{
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
		public Connect()
		{
		}

		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_application = (DTE2)application;
			_addinInstance = (AddIn)addInInst;

			CommandBar menuBarCommandBar = ((CommandBars)_application.CommandBars)["MenuBar"];
			CommandBarControl toolsControl = menuBarCommandBar.Controls["Tools"];
			CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;
			CommandBarPopup boilerplatePopup = null;

			try
			{
				boilerplatePopup = (CommandBarPopup)toolsPopup.Controls["Boilerplate"];
				boilerplatePopup.Delete(false);
				boilerplatePopup = null;
			}
			catch
			{
			}

			if ((connectMode == ext_ConnectMode.ext_cm_Startup || connectMode == ext_ConnectMode.ext_cm_AfterStartup) &&
				(boilerplatePopup == null))
			{
				boilerplatePopup = (CommandBarPopup)toolsPopup.Controls.Add(MsoControlType.msoControlPopup);
				boilerplatePopup.Caption = "Boilerplate";
				boilerplatePopup.BeginGroup = true;

				Command autoCommand = _application.Commands.AddNamedCommand(_addinInstance, "AutoFile", "Auto", "Generate appropriate boilerplate code.", true);
				CommandBarControl autoControl = (CommandBarControl)autoCommand.AddControl(boilerplatePopup.CommandBar, 1);
				autoControl.Caption = "Automatic";

				Command headerCommand = _application.Commands.AddNamedCommand(_addinInstance, "HeaderFile", "Header", "Generate header file boilerplate code.", true);
				CommandBarControl headerControl = (CommandBarControl)headerCommand.AddControl(boilerplatePopup.CommandBar, 2);
				headerControl.Caption = "Generate Header File";

				Command implemenatationCommand = _application.Commands.AddNamedCommand(_addinInstance, "ImplementationFile", "Impl", "Generate implementation file boilerplate code.", true);
				CommandBarControl implControl = (CommandBarControl)implemenatationCommand.AddControl(boilerplatePopup.CommandBar, 3);
				implControl.Caption = "Generate Implementation File";
			}
		}

		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		public void OnAddInsUpdate(ref Array custom)
		{
		}

		public void OnStartupComplete(ref Array custom)
		{
		}

		public void OnBeginShutdown(ref Array custom)
		{
		}

		public void Exec(string CmdName, vsCommandExecOption ExecuteOption, ref object VariantIn, ref object VariantOut, ref bool Handled)
		{
			Handled = false;
			Document currentDocument = _application.ActiveDocument;

			if (ExecuteOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
				if (CmdName == "Boilerplate.Connect.HeaderFile")
				{
					InsertHeaderFileBoilerplate(currentDocument);
					Handled = true;
				}
				else if (CmdName == "Boilerplate.Connect.ImplementationFile")
				{
					InsertImplementationFileBoilerplate(currentDocument);
					Handled = true;
				}
				else if (CmdName == "Boilerplate.Connect.AutoFile")
				{
					if (IsHeaderFile(currentDocument))
					{
						InsertHeaderFileBoilerplate(currentDocument);
						Handled = true;
					}
					else if (IsImplementationFile(currentDocument))
					{
						InsertImplementationFileBoilerplate(currentDocument);
						Handled = true;
					}
				}
			}
		}

		public void QueryStatus(string CmdName, vsCommandStatusTextWanted NeededText, ref vsCommandStatus StatusOption, ref object CommandText)
		{
			Document currentDocument = _application.ActiveDocument;
			StatusOption = vsCommandStatus.vsCommandStatusUnsupported;
			
			if (NeededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
				if (CmdName == _addinInstance.ProgID + ".HeaderFile")
				{
					StatusOption = vsCommandStatus.vsCommandStatusEnabled | vsCommandStatus.vsCommandStatusSupported;
				}
				else if (CmdName == _addinInstance.ProgID + ".ImplementationFile")
				{
					StatusOption = vsCommandStatus.vsCommandStatusEnabled | vsCommandStatus.vsCommandStatusSupported;
				}
				else if (CmdName == _addinInstance.ProgID + ".AutoFile")
				{
					StatusOption = vsCommandStatus.vsCommandStatusEnabled | vsCommandStatus.vsCommandStatusSupported;
				}
			}
		}

		private string[] _headerFileExtensions = { ".h", ".hh", ".hpp", ".hxx" };
		private string[] _implementationFileExtensions = { ".c", ".cc", ".cpp", ".cxx" };

		private bool IsHeaderFile(Document document)
		{
			String extension = Path.GetExtension(document.FullName);
			if (String.IsNullOrEmpty(extension)) return false;
			return Array.Exists(_headerFileExtensions, e => e == extension);
		}

		private bool IsImplementationFile(Document document)
		{
			String extension = Path.GetExtension(document.FullName);
			if (String.IsNullOrEmpty(extension)) return false;
			return Array.Exists(_implementationFileExtensions, e => e == extension);
		}

		private class UndoBlock : IDisposable
		{
			public UndoBlock(DTE2 application, String description)
			{
				_context = null;
				if (!application.UndoContext.IsOpen)
				{
					_context = application.UndoContext;
					_context.Open(description);
				}
			}

			public void Dispose()
			{
				if (_context != null)
				{
					_context.Close();
				}
			}

			private UndoContext _context;
		}

		private void InsertHeaderFileBoilerplate(Document document)
		{
			String filename = Path.GetFileName(document.FullName);
			String projectname = document.ProjectItem.ContainingProject.Name;
			String header_guard = "__" + projectname.ToUpper() + "_" + filename.Replace(".", "_").ToUpper() + "__";

			TextSelection selection = (TextSelection)document.Selection;
			EditPoint editPoint = selection.ActivePoint.CreateEditPoint();

			using (new UndoBlock(_application, "Inserted header file boilerplate code"))
			{
				editPoint.StartOfDocument();
				editPoint.Insert("#pragma once\n");
				editPoint.Insert("#ifndef " + header_guard + "\n");
				editPoint.Insert("#define " + header_guard + "\n");
				editPoint.Insert("\n");
				EditPoint cursorPoint = editPoint.CreateEditPoint();
				editPoint.Insert("\n");

				editPoint.EndOfDocument();
				editPoint.Insert("\n#endif // " + header_guard + "\n");

				selection.GotoLine(cursorPoint.Line);
			}
		}

		private void InsertImplementationFileBoilerplate(Document document)
		{
			String basename = Path.GetFileNameWithoutExtension(document.FullName);
			String extension = Path.GetExtension(document.FullName);
			String header_extension = extension.Replace("c", "h");

			TextSelection selection = (TextSelection)document.Selection;
			EditPoint editPoint = selection.ActivePoint.CreateEditPoint();

			String precompiled_header = "Precompiled";
			try
			{
				precompiled_header = (String)document.ProjectItem.Properties.Item("Precompiled Header File").Value;
			}
			catch (Exception e)
			{
			}

			using (new UndoBlock(_application, "Inserted implementation file boilerplate code"))
			{
				editPoint.StartOfDocument();
				editPoint.Insert("#include \"" + precompiled_header + header_extension + "\"\n");
				editPoint.Insert("#include \"" + basename + header_extension + "\"\n");
				editPoint.Insert("\n");

				selection.GotoLine(editPoint.Line);
			}
		}

		private DTE2 _application;
		private AddIn _addinInstance;

	}
}