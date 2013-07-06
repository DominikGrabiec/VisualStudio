using System;
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

		private CommandBarPopup boilerplatePopup = null;

		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_application = (DTE2)application;
			_addinInstance = (AddIn)addInInst;

			Command autoCommand = null;
			Command headerCommand = null;
			Command implemenatationCommand = null;

			if (connectMode != ext_ConnectMode.ext_cm_UISetup)
			{
				try
				{
					autoCommand = _application.Commands.Item(_addinInstance.ProgID + ".AutoFile");
				}
				catch
				{
				}
				try
				{
					headerCommand = _application.Commands.Item(_addinInstance.ProgID + ".HeaderFile");
				}
				catch
				{
				}
				try
				{
					implemenatationCommand = _application.Commands.Item(_addinInstance.ProgID + ".ImplementationFile");
				}
				catch
				{
				}
			}
			
			if (autoCommand == null)
			{
				autoCommand = _application.Commands.AddNamedCommand(_addinInstance, "AutoFile", "Auto", "Generate appropriate boilerplate code.", true);
			}
			if (headerCommand == null)
			{
				headerCommand = _application.Commands.AddNamedCommand(_addinInstance, "HeaderFile", "Header", "Generate header file boilerplate code.", true);
			}
			if (implemenatationCommand == null)
			{
				implemenatationCommand = _application.Commands.AddNamedCommand(_addinInstance, "ImplementationFile", "Impl", "Generate implementation file boilerplate code.", true);
			}

			CommandBar menuBarCommandBar = ((CommandBars)_application.CommandBars)["MenuBar"];
			CommandBarControl toolsControl = menuBarCommandBar.Controls["Tools"];
			CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

			try
			{
				boilerplatePopup = (CommandBarPopup)toolsPopup.Controls["Boilerplate"];
			}
			catch
			{
			}

			if (boilerplatePopup == null)
			{
				boilerplatePopup = (CommandBarPopup)toolsPopup.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, toolsPopup.Controls.Count + 1, true);
				boilerplatePopup.CommandBar.Name = "Boilerplate";
				boilerplatePopup.Caption = "Boilerplate";
				boilerplatePopup.BeginGroup = true;

				CommandBarControl autoControl = (CommandBarControl)autoCommand.AddControl(boilerplatePopup.CommandBar, 1);
				autoControl.Caption = "Automatic";

				CommandBarControl headerControl = (CommandBarControl)headerCommand.AddControl(boilerplatePopup.CommandBar, 2);
				headerControl.Caption = "Generate Header File";

				CommandBarControl implControl = (CommandBarControl)implemenatationCommand.AddControl(boilerplatePopup.CommandBar, 3);
				implControl.Caption = "Generate Implementation File";
			}

		}

		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
			try
			{
				if (boilerplatePopup != null)
				{
					boilerplatePopup.Delete(true);
				}
			}
			catch
			{
			}
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
				if (CmdName == _addinInstance.ProgID + ".HeaderFile")
				{
					InsertHeaderFileBoilerplate(currentDocument);
					Handled = true;
				}
				else if (CmdName == _addinInstance.ProgID + ".ImplementationFile")
				{
					InsertImplementationFileBoilerplate(currentDocument);
					Handled = true;
				}
				else if (CmdName == _addinInstance.ProgID + ".AutoFile")
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
				editPoint.Insert("#pragma once\r\n");
				editPoint.Insert("#ifndef " + header_guard + "\r\n");
				editPoint.Insert("#define " + header_guard + "\r\n");
				editPoint.Insert("\r\n");
				EditPoint cursorPoint = editPoint.CreateEditPoint();
				editPoint.Insert("\r\n");

				editPoint.EndOfDocument();
				editPoint.Insert("\r\n#endif // " + header_guard + "\r\n");

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

			String precompiled_header = "Precompiled.hpp";
			try
			{
				precompiled_header = (String)document.ProjectItem.Properties.Item("Precompiled Header File").Value;
			}
			catch
			{
			}

			using (new UndoBlock(_application, "Inserted implementation file boilerplate code"))
			{
				editPoint.StartOfDocument();
				editPoint.Insert("#include \"" + precompiled_header + "\"\r\n");
				editPoint.Insert("#include \"" + basename + header_extension + "\"\r\n");
				editPoint.Insert("\r\n");

				selection.GotoLine(editPoint.Line);
			}
		}

		private DTE2 _application;
		private AddIn _addinInstance;

	}
}