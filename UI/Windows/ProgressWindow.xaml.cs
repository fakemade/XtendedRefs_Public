using System.Windows;

namespace XtendedRefs.ComAddIn.UI.Windows;

public partial class ProgressWindow : Window
{
    private int _current;
    private int _total;

    public ProgressWindow()
    {
        InitializeComponent();
    }

    public ProgressWindow(string title, int max) : this()
    {
        Title = title;
        _total = max;
        ProgressBar.Maximum = max;
        StatusLabel.Text = $"0 / {max}";
    }

    public bool CancelRequested { get; private set; }

    public void SetMax(int max)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.Invoke(() => SetMax(max));
            return;
        }

        _total = max;
        ProgressBar.Maximum = max;
        StatusLabel.Text = $"{_current} / {max}";
    }

    public void Increment(string message = "")
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.Invoke(() => Increment(message));
            return;
        }

        _current++;
        int value = _current > _total ? _total : _current;
        ProgressBar.Value = value;
        StatusLabel.Text = string.IsNullOrEmpty(message)
            ? $"{_current} / {_total}"
            : $"{_current} / {_total} — {message}";
    }

    public void SetMessage(string message)
    {
        if (!Dispatcher.CheckAccess())
        {
            Dispatcher.Invoke(() => SetMessage(message));
            return;
        }

        StatusLabel.Text = string.IsNullOrEmpty(message)
            ? $"{_current} / {_total}"
            : $"{_current} / {_total} — {message}";
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        CancelRequested = true;
        DialogResult = false;
        Close();
    }
}
