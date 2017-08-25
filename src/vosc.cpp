#include "vosc.h"
#include <QDebug>
#include <QMessageBox>
#include <Windows.h>

//class VOSC

VOSC::VOSC(QWidget *parent) : QWidget(parent)
{
    qDebug()<<"VOSC";
    time = QDateTime::currentDateTime();
}

void VOSC::init(int count,int capacity, int max, int min, int dev_grid, int interval)
{
    if((max - min)%dev_grid != 0)
        qDebug()<<"[[WARNING]] __ Cannot allocate height properly!";
    layout_main = new QHBoxLayout(this);
    layout_options = new QVBoxLayout();
    widget_osc = new WidgetOsc(this);
    list_element = new QList<Element*>();
    list_curve = new QList<QList<float>*>();
    pause_button = new QPushButton();
    reset_button = new QPushButton();
    print_button = new QPushButton();
    save_as_excel_button = new QPushButton();

    curve_count = count;
    capacity_value = capacity;
    max_value = max;
    min_value = min;
    grid_dev_value = dev_grid;
    interval_value = interval;

    initOption();
    initCurve();
    initOsc(getMaxDigit(max_value, min_value));
}

void VOSC::initOption()
{
    layout_main->addWidget(widget_osc);
    layout_main->addLayout(layout_options);
    layout_main->setStretch(0,5);
    layout_main->setStretch(1,1);
    for(int i = 0 ; i < curve_count; i++){
        Element* new_element = new Element();
        list_element->append(new_element);
        layout_options->addLayout(new_element);
    }
    layout_options->addStretch();

    print_button->setText("Print");
    reset_button->setText("Reset");
    save_as_excel_button->setText("Save as Excel");
    QHBoxLayout* pause_print_button_layout = new QHBoxLayout();
    pause_print_button_layout->addWidget(pause_button);
    pause_print_button_layout->addWidget(print_button);
    layout_options->addLayout(pause_print_button_layout);
    layout_options->addWidget(reset_button);
    layout_options->addWidget(save_as_excel_button);
    this->setLayout(layout_main);

    pause_button->setText("Pause");

    connect(pause_button,&QPushButton::clicked,[=](){
        if(pause_button->text() == "Paused")
            pause_button->setText("Pause");
        else if(pause_button->text() == "Pause")
            pause_button->setText("Paused");
    });
    connect(reset_button,&QPushButton::clicked,[=](){
        widget_osc->row_height -= widget_osc->mouse_wheel_count*3;
        widget_osc->origin_position -= widget_osc->mouse_move_count*3.5;
        widget_osc->mouse_wheel_count = 0;
        widget_osc->mouse_move_count = 0;
    });
    connect(print_button,&QPushButton::clicked,[=](){
        widget_osc->saveImage();
    });
    connect(save_as_excel_button,&QPushButton::clicked,this,&VOSC::saveExcel);
}

void VOSC::initCurve()
{
    for(int i = 0; i < curve_count; i++){
        auto new_curve = new QList<float>();
        list_curve->append(new_curve);
    }
}

void VOSC::initOsc(int digit)
{
    widget_osc->init(digit, (max_value - min_value)/grid_dev_value);
}

void VOSC::setCurveAttribute(int index, QString name, QColor& color)
{
    list_element->at(index)->setColor(color);
    list_element->at(index)->check->setText(name);
}

int VOSC::getMaxDigit(int max, int min)
{
    int tempMax = max;
    int tempMin = abs(min);
    int countMax = 0;
    int countMin = 0;
    while(tempMax > 0){
        countMax++;
        tempMax /= 10;
    }
    while(tempMin > 0){
        countMin++;
        tempMin /= 10;
    }
    if(min < 0)
        countMin++;
    return (countMax > countMin ? countMax : countMin);
}

void VOSC::addPoint(int index, float value)
{
    if(pause_button->text() == "Paused")
        return;
    emit SaveAsExcelSignal(index, value);
    if(index == curve_count - 1){
        QDateTime time2 = QDateTime::currentDateTime();
        //qDebug()<<1/(time.msecsTo(time2) / 1000.0);
        time = QDateTime::currentDateTime();
    }
    list_curve->at(index)->append(value);
    list_element->at(index)->edit->setText(QString::number(value));
    if(list_curve->at(curve_count - 1)->count() > capacity_value){
        if(index == curve_count - 1)
            widget_osc->move_count = (widget_osc->move_count+1) % (this->interval_value - 1);
        for(int i = 0;i < curve_count; i ++){
            if(list_curve->at(i)->isEmpty())
                continue;
            list_curve->at(i)->pop_front();
        }
    }

}

void VOSC::drawCurve()
{
    if(pause_button->text() == "Pause")
        widget_osc->drawCurve();
}

void VOSC::saveExcel()
{

    if(save_as_excel_button->text() == "Save as Excel"){
        save_as_excel_button->setText("Recording..");
        save_as_excel_thread = new SaveAsExcelThread(this);
        save_as_excel_thread->start();
    }
    else if(save_as_excel_button->text() == "Recording.."){
        save_as_excel_button->setText("Save as Excel");
        save_as_excel_thread->quit();
    }
}

int VOSC::getCurveCount()
{
    return curve_count;
}

//class VOSC::WidgetOsc

VOSC::WidgetOsc::WidgetOsc(VOSC* osc,QWidget* parent) :
    QWidget(parent), current_osc(osc)
{

}

void VOSC::WidgetOsc::init(int _start, int _count)
{
    start = _start*7;
    row_count = _count;
    move_count = 0;
    is_midbutton_pressed = false;
    current_mouse_y = 0;
    mouse_wheel_count = 0;
    mouse_move_count = 0;
    col_interval = (this->width()-start)/(double)(current_osc->capacity_value - 1);
}

void VOSC::WidgetOsc::drawCurve()
{
    QWidget::update();
}

void VOSC::WidgetOsc::drawCurve(QPaintDevice *device)
{
    QPainter painter(device);
    painter.setRenderHint(QPainter::Antialiasing);
    QPen pen;
    for(int i = 0; i < current_osc->curve_count; i++){
        pen.setColor(current_osc->list_element->at(i)->color);
        painter.setPen(pen);
        for(int j = 0; j < current_osc->list_curve->at(i)->count() - 1; j ++){
            if(!current_osc->list_element->at(i)->check->isChecked())
                continue;
            painter.drawLine(j*col_interval + start,
                       origin_position - current_osc->list_curve->at(i)->at(j)*row_height/(double)current_osc->grid_dev_value,
                       (j+1)*col_interval + start,
                       origin_position - current_osc->list_curve->at(i)->at(j+1)*row_height/(double)current_osc->grid_dev_value);
            painter.drawEllipse(QPoint((j+1)*col_interval + start,
                                       origin_position - current_osc->list_curve->at(i)->at(j+1)*row_height/(double)current_osc->grid_dev_value),2,2);
        }
    }
}

void VOSC::WidgetOsc::paintEvent(QPaintEvent* event)
{
    drawGrid(this);
    drawCurve(this);
}

void VOSC::WidgetOsc::resizeEvent(QResizeEvent* event)
{
    row_height = (this->height() - 10)/(double)row_count;
    col_interval = (this->width()-start)/(double)(current_osc->capacity_value - 1);
    origin_position = (current_osc->max_value/current_osc->grid_dev_value)*row_height + 5;
}

void VOSC::WidgetOsc::wheelEvent(QWheelEvent *event)
{
    current_row_count = this->height()/row_height;
    if(event->delta() > 0 && current_row_count > 3){
        row_height+=3.00;
        mouse_wheel_count++;
    }
    if(event->delta() < 0 && current_row_count < row_count*2.5){
        row_height-=3.00;
        mouse_wheel_count--;
    }
}

void VOSC::WidgetOsc::mousePressEvent(QMouseEvent *event)
{
    if(event->button() == Qt::MidButton){
        current_mouse_y = event->globalY();
        is_midbutton_pressed = true;
    }
}

void VOSC::WidgetOsc::mouseReleaseEvent(QMouseEvent *event)
{
    if(event->button() == Qt::MidButton && is_midbutton_pressed == true){
        is_midbutton_pressed = false;
    }
}

void VOSC::WidgetOsc::mouseMoveEvent(QMouseEvent *event)
{
    if(is_midbutton_pressed){
        int temp_mouse_y = event->globalY();
        if(temp_mouse_y > current_mouse_y){
            origin_position+=3.50;
            mouse_move_count++;
        }
        else if(temp_mouse_y < current_mouse_y){
            origin_position-=3.50;
            mouse_move_count--;
        }
        current_mouse_y = temp_mouse_y;
    }
}

void VOSC::WidgetOsc::drawGrid(QPaintDevice* device)
{
    current_row_count = 0;
    QPainter painter(device);
    QPen pen;
    pen.setColor(Qt::gray);
    painter.setPen(pen);

    /*************************************************************
     * --old function--
     * --draw lines from row0--
     *
    int val = current_osc->max_value;
    for(double i = 5; i <= device->height(); i+=row_height){
        painter.drawLine(start,i,device->width(),i);
        painter.drawText(QPoint(0,i+5), QString::number(val));
        val -= current_osc->grid_dev_value;
    }

    **************************************************************/

    int val = 0;
    for(double i = origin_position; i <= device->height(); i+=row_height){
        painter.drawLine(start,i,device->width(),i);
        painter.drawText(QPoint(0,i+5), QString::number(val));
        val -= current_osc->grid_dev_value;
    }
    val = 0;
    for(double i = origin_position; i >= 4; i-=row_height){
        painter.drawLine(start,i,device->width(),i);
        painter.drawText(QPoint(0,i+5), QString::number(val));
        val += current_osc->grid_dev_value;
    }

    painter.drawLine(start,5,start,device->height() - 6);
    painter.drawLine(device->width() - 1,5,device->width() - 1,device->height() - 6);
    for(double i = start + col_interval*(current_osc->interval_value - 1) - move_count*col_interval; i < device->width(); i+=col_interval*(current_osc->interval_value - 1)){
        painter.drawLine(i,5,i,device->height() - 6);
    }
}

void VOSC::WidgetOsc::saveImage()
{
    QPixmap* capture = new QPixmap(this->size());
    capture->fill();
    SaveImageThread* t = new SaveImageThread(this);
    t->start();
}

//class VOSC::SaveImageThread

VOSC::SaveImageThread::SaveImageThread(WidgetOsc *_osc)
{
    osc = _osc;

    connect(this,&SaveImageThread::finished,this,&SaveImageThread::deleteLater);
}

void VOSC::SaveImageThread::run()
{
    QPixmap* img = new QPixmap(osc->size());
    img->fill();
    osc->drawGrid(img);
    osc->drawCurve(img);
    QString image_name = "F:\\JUSTZONE_GroundStation\\Capture\\capture_" + osc->current_osc->objectName() + QDateTime::currentDateTime().toString("_MMddhhmmss") + ".png";
    img->save(image_name);
    delete img;
}

//class VOSC::SaveAsExcelThread

SaveAsExcelThread::SaveAsExcelThread(VOSC* _osc)
{
    is_quit = false;
    osc = _osc;

    connect(osc,&VOSC::SaveAsExcelSignal,this,&SaveAsExcelThread::SaveAsExcel);
    connect(this,&SaveAsExcelThread::finished,this,&SaveAsExcelThread::deleteLater);
}

void SaveAsExcelThread::run()
{
    count = 0;
    list = new QList<QList<QVariant>*>();
    QString fileName = "F:\\JUSTZONE_GroundStation\\Excels\\" + osc->objectName() + QDateTime::currentDateTime().toString("_MMddhhmmss") + ".xlsx";

    qDebug()<<"start";

    while(1)
        if(is_quit && count%(osc->getCurveCount()) == 0)
            break;

    qDebug()<<"quit!";

    for(int i = 0; i < list->size(); i++){
        for(int j = 0; j < list->at(0)->size(); j++)
            qDebug()<<list->at(i)->at(j).toFloat();
        qDebug()<<"*********";
    }

    QVariantList varList;
    for(int i = 0; i < list->size(); i++)
        varList.append(QVariant(*(list->at(i))));
    QVariant variant(varList);

    QString rangeStr;
    CastToWord(list->at(0)->size(),rangeStr);
    rangeStr = "A1:"+rangeStr+QString::number(list->size());
    qDebug()<<rangeStr;

    HRESULT r = OleInitialize(0);

    qDebug()<<"writing..";

    QAxObject* excel = new QAxObject(this);
    excel->setControl("Excel.Application");
    excel->dynamicCall("SetVisible(bool Visible)",false);
    excel->setProperty("DisplayAlerts",false);
    QAxObject* workbooks = excel->querySubObject("WorkBooks");
    workbooks->dynamicCall("Add");
    QAxObject* workbook = excel->querySubObject("ActiveWorkBook");
    QAxObject* worksheet = workbook->querySubObject("WorkSheets(int)",1);
    workbook->dynamicCall("SaveAs(const QString&)",fileName);
    QAxObject* range = worksheet->querySubObject("Range(const QString&)",rangeStr);
    range->setProperty("Value",variant);
    workbook->dynamicCall("Save()",fileName);
    workbook->dynamicCall("Close(Boolean)",false);

    qDebug()<<"done";

    OleUninitialize();

    delete list;
    list = NULL;
}

void SaveAsExcelThread::quit()
{
    is_quit = true;
}

void SaveAsExcelThread::CastToWord(int data, QString &res)
{
    Q_ASSERT(data>0 && data<65535);
    int tempData = data / 26;
    if(tempData > 0){
        int mode = data % 26;
        CastToWord(mode,res);
        CastToWord(tempData,res);
    }
     else{
        res = QString(QChar(data + 0x40)) + res;
    }
}

void SaveAsExcelThread::SaveAsExcel(int index, float value)
{
    if(is_quit && count%(osc->getCurveCount()) == 0)
        return;
    if(count%(osc->getCurveCount()) == 0)
        list->append(new QList<QVariant>());
    list->at(count/(osc->getCurveCount()))->append(value);
    //qDebug()<<osc->objectName()<<" -- "<<list->size()<<" -- "<<index<<" -- "<<value;
    count++;
}

//class Element

Element::Element(QWidget* parent, QString& name, bool _is_check, QColor& _color) : QHBoxLayout(parent)
{
    is_check = _is_check;
    check = new QCheckBox(name);
    edit = new QLineEdit();
    button = new ColorButton();
    color = _color;

    initElement();
}

void Element::initElement()
{
    check->setChecked(is_check);
    this->addWidget(check);
    this->addWidget(edit);
    this->addWidget(button);
    this->setSpacing(3);

    QString style = "background-color: rgb(%1,%2,%3)";
    button->widget->setStyleSheet(style.arg(color.red()).arg(color.green()).arg(color.blue()));

    connect(button,&QPushButton::clicked,this,&Element::OnButtonClicked);
}

void Element::setColor(QColor& _color)
{
    color = _color;
    QString style = "background-color: rgb(%1,%2,%3)";
    button->widget->setStyleSheet(style.arg(color.red()).arg(color.green()).arg(color.blue()));
}

void Element::OnButtonClicked()
{
    QColor temp_color = QColorDialog::getColor(color);
    bool flag = temp_color.isValid();
    qDebug()<<"flag = "<<flag;
    if(flag){
        setColor(temp_color);
        color = temp_color;
    }
}

//class ColorButton

ColorButton::ColorButton(QWidget *parent) : QPushButton(parent)
{
    QHBoxLayout* lyt = new QHBoxLayout();
    widget = new QWidget();
    lyt->addWidget(widget);
    lyt->setMargin(5);
    this->setLayout(lyt);
}
