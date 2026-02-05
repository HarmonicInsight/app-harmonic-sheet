using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using HarmonicSheet.Services;

namespace HarmonicSheet.Views;

public partial class TutorialWindow : Window
{
    private readonly ITutorialService _tutorialService;

    public TutorialWindow(ITutorialService tutorialService)
    {
        InitializeComponent();
        _tutorialService = tutorialService;

        CreateProgressDots();
        UpdateContent();
    }

    private void CreateProgressDots()
    {
        ProgressDots.Children.Clear();

        for (int i = 0; i < _tutorialService.TotalSteps; i++)
        {
            var dot = new Ellipse
            {
                Width = 12,
                Height = 12,
                Margin = new Thickness(4, 0, 4, 0),
                Fill = new SolidColorBrush(Color.FromRgb(203, 213, 225)) // gray
            };
            ProgressDots.Children.Add(dot);
        }
    }

    private void UpdateProgressDots()
    {
        for (int i = 0; i < ProgressDots.Children.Count; i++)
        {
            if (ProgressDots.Children[i] is Ellipse dot)
            {
                dot.Fill = i == _tutorialService.CurrentStep
                    ? new SolidColorBrush(Color.FromRgb(37, 99, 235))   // blue (current)
                    : i < _tutorialService.CurrentStep
                        ? new SolidColorBrush(Color.FromRgb(34, 197, 94))  // green (completed)
                        : new SolidColorBrush(Color.FromRgb(203, 213, 225)); // gray (pending)
            }
        }
    }

    private void UpdateContent()
    {
        var step = _tutorialService.GetCurrentStep();

        if (step == null || _tutorialService.IsCompleted)
        {
            DialogResult = true;
            Close();
            return;
        }

        StepIndicator.Text = $"ステップ {step.StepNumber} / {_tutorialService.TotalSteps}";
        TitleText.Text = step.Title;
        DescriptionText.Text = step.Description;

        // ボタンの状態を更新
        PrevButton.Visibility = _tutorialService.CurrentStep > 0
            ? Visibility.Visible
            : Visibility.Collapsed;

        // 最後のステップでは「次へ」を「始める」に変更
        if (_tutorialService.CurrentStep == _tutorialService.TotalSteps - 1)
        {
            NextButton.Content = "始める！";
            NextButton.Background = new SolidColorBrush(Color.FromRgb(5, 150, 105)); // green
        }
        else
        {
            NextButton.Content = "次へ →";
            NextButton.Background = new SolidColorBrush(Color.FromRgb(37, 99, 235)); // blue
        }

        UpdateProgressDots();
    }

    private void OnNextClick(object sender, RoutedEventArgs e)
    {
        _tutorialService.NextStep();
        UpdateContent();
    }

    private void OnPrevClick(object sender, RoutedEventArgs e)
    {
        _tutorialService.PreviousStep();
        UpdateContent();
    }

    private void OnSkipClick(object sender, RoutedEventArgs e)
    {
        var result = MessageBox.Show(
            "チュートリアルをスキップしますか？\n\n後から「ヘルプ」→「使い方ガイド」で\nいつでも見ることができます。",
            "確認",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Yes)
        {
            _tutorialService.Skip();
            DialogResult = true;
            Close();
        }
    }
}
