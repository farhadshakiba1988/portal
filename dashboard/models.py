from django.db import models

class Announcement(models.Model):
    """مدل اطلاعیه‌ها"""
    tp_ID = models.AutoField(primary_key=True)
    nvarchar1 = models.CharField(max_length=255, verbose_name="عنوان")
    ntext2 = models.TextField(verbose_name="متن")
    tp_Created = models.DateTimeField(auto_now_add=True, verbose_name="تاریخ ایجاد")
    tp_Modified = models.DateTimeField(auto_now=True, verbose_name="تاریخ ویرایش")

    class Meta:
        db_table = "Announcements"  # نام جدول مشابه SharePoint
        verbose_name = "اطلاعیه"
        verbose_name_plural = "اطلاعیه‌ها"

    def __str__(self):
        return self.nvarchar1
