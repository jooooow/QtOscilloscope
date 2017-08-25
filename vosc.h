#ifndef VOSC_H
#define VOSC_H

#include <QObject>
#include <QWidget>
#include <QLayout>
#include <QPushButton>
#include <QCheckBox>
#include <QLineEdit>
#include <QSpacerItem>
#include <QPainter>
#include <QPen>
#include <QTimer>
#include <QColorDialog>
#include <QPalette>
#include <QColor>
#include <QDateTime>
#include <QThread>
#include <QAxObject>
#include <QVariant>

#include <QPaintEvent>
#include <QResizeEvent>
#include <QWheelEvent>
#include <QMouseEvent>

class Element;
class ColorButton;
class SaveAsExcelThread;

class VOSC : public QWidget
{
    Q_OBJECT
public:
    explicit VOSC(QWidget *parent = 0);
    void init(int count = 3,
              int capacity = 100,
              int max = 100,
              int min = -100,
              int dev_grid = 10,
              int interval = 10
            );
    void setCurveAttribute(int index, QString name = "name", QColor& color = QColor(0,0,196));
    void addPoint(int index, float value);
    void drawCurve();
    void saveExcel();
    int getCurveCount();

private:
    void initOption();
    void initCurve();
    void initOsc(int digit);
    int getMaxDigit(int max, int min);

private:
    //nested class
    class WidgetOsc : public QWidget
    {
    public:
        explicit WidgetOsc(VOSC* osc,QWidget* parent = nullptr);
        void init(int _start, int _count);
        void drawCurve();
        void drawCurve(QPaintDevice* device);
        void drawGrid(QPaintDevice* device);
        void saveImage();

    protected:
        void paintEvent(QPaintEvent* event);
        void resizeEvent(QResizeEvent* event);
        void wheelEvent(QWheelEvent* event);
        void mousePressEvent(QMouseEvent* event);
        void mouseReleaseEvent(QMouseEvent* event);
        void mouseMoveEvent(QMouseEvent* event);

    public:
        bool is_midbutton_pressed;
        int start;
        int row_count;
        int move_count;
        int current_row_count;
        int current_mouse_y;
        int mouse_wheel_count;
        int mouse_move_count;
        double origin_position;
        double row_height;
        double col_interval;
        const VOSC* current_osc;
    };
    class SaveImageThread : public QThread
    {
    public:
        explicit SaveImageThread(WidgetOsc* _osc);

    protected:
        void run();

    private:
        WidgetOsc* osc;
    };

private:
    QHBoxLayout* layout_main;
    QVBoxLayout* layout_options;
    WidgetOsc* widget_osc;
    QList<Element*>* list_element;
    QList<QList<float>*>* list_curve;
    QDateTime time;
    QPushButton* pause_button;
    QPushButton* print_button;
    QPushButton* reset_button;
    QPushButton* save_as_excel_button;
    SaveAsExcelThread* save_as_excel_thread;

private:
    int curve_count;
    int capacity_value;
    int max_value;
    int min_value;
    int grid_dev_value;
    int interval_value;
    int frequency;

signals:
    void SaveAsExcelSignal(int, float);
public slots:
};

class Element : public QHBoxLayout
{
    Q_OBJECT
public:
    explicit Element(QWidget* parent = nullptr,
                     QString& name = QString("check"),
                     bool _is_check = true,
                     QColor& _color = QColor(0,0,196));
    void initElement();
    void setColor(QColor& _color);

public:
    bool is_check;
    QCheckBox* check;
    QLineEdit* edit;
    ColorButton* button;
    QColor color;

private slots:
    void OnButtonClicked();
};

class ColorButton : public QPushButton
{
public :
    explicit ColorButton(QWidget* parent = nullptr);
    QWidget* widget;
};

class SaveAsExcelThread : public QThread
{
    Q_OBJECT
public:
    explicit SaveAsExcelThread(VOSC* _osc);
    void quit();

protected:
    void run();

private:
    void SaveAsExcel(int index, float value);
    void CastToWord(int data, QString& res);

private:
    bool is_quit;
    int count;
    VOSC* osc;
    QList<QList<QVariant>*>* list;
};

#endif // VOSC_H
