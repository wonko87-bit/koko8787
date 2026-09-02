using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Input;
using Flowdeck.Windows.ViewModels;
using Microsoft.Win32;

namespace Flowdeck.Windows.Views;

/// <summary>
/// The weekly report draft: a week of the workspace as text, in a box that can be edited
/// before it is copied out. Reached from the tray menu — it belongs to no list window.
/// </summary>
public partial class ReportWindow : Window
{
    private readonly ReportViewModel _viewModel;
    private readonly string _folder;

    public ReportWindow(ReportViewModel viewModel, string folder)
    {
        InitializeComponent();

        _viewModel = viewModel;
        _folder = folder;
        DataContext = viewModel;
    }

    /// <summary>
    /// The whole draft onto the clipboard, which is how it gets to wherever the report is
    /// actually filed. Somebody else may be holding the clipboard for a moment; that is a
    /// line of status, not a crash.
    /// </summary>
    private void OnCopyClick(object sender, RoutedEventArgs e)
    {
        try
        {
            Clipboard.SetText(_viewModel.Text);
            _viewModel.Status = "복사했습니다. 붙여넣을 곳에서 Ctrl+V 하세요";
        }
        catch (Exception ex) when (ex is ExternalException or InvalidOperationException)
        {
            _viewModel.Status = "클립보드를 쓰지 못했습니다. 잠시 후 다시 시도해 주세요";
        }
    }

    /// <summary>
    /// Saves the draft as it stands, edits and all. Defaults into a folder of its own under
    /// the data folder, named by the week, so a year of them sorts by itself.
    /// </summary>
    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        var reports = Path.Combine(_folder, "reports");

        var dialog = new SaveFileDialog
        {
            Title = "주간보고 저장",
            FileName = _viewModel.SuggestedFileName,
            DefaultExt = ".txt",
            Filter = "텍스트 파일 (*.txt)|*.txt|모든 파일 (*.*)|*.*",
            InitialDirectory = Directory.Exists(reports) ? reports : _folder,
        };

        if (dialog.ShowDialog(this) != true) return;

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(dialog.FileName) ?? _folder);
            File.WriteAllText(dialog.FileName, _viewModel.Text, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            _viewModel.Status = "저장했습니다: " + dialog.FileName;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _viewModel.Status = "저장하지 못했습니다: " + ex.Message;
        }
    }

    private void OnHeaderDrag(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton != MouseButton.Left) return;

        try
        {
            DragMove();
        }
        catch (InvalidOperationException)
        {
        }
    }

    private void OnCloseClick(object sender, RoutedEventArgs e) => Close();
}
