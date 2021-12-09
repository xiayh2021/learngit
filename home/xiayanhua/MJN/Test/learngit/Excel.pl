#!/usr/local/bin/perl -w
# 
# Copyright (c) BMK 2011
my $ver="1.0";
# Writer:         Xiayh <xiayh@biomarker.com.cn>
# Program Date:   2011.
# Modifier:       Xiayh <xiayh@biomarker.com.cn>
use strict;
use Cwd;
use Getopt::Long;
use Data::Dumper;
use FindBin qw($Bin $Script);
use File::Basename qw(basename dirname);
use Spreadsheet::ParseExcel;  ##  这种模块 只能读取 xls的
use Spreadsheet::ParseXLSX;   ## 这种模块 可以读取xlsx的
use Excel::Writer::XLSX;   ## 20211209 
use Spreadsheet::ParseExcel::FmtUnicode;
######################请在写程序之前，一定写明时间、程序用途、参数说明；每次修改程序时，也请做好注释工作

my %opts;
GetOptions(\%opts,"i=s","o=s","h" );

#&help()if(defined $opts{h});
if(!defined($opts{i})     ||!defined($opts{o})|| defined($opts{h}))
{
	print <<"	Usage End.";
	Description:
	#在PE文件里面提取出来read1  和 read2的信息，得到的是fa格式的，查看kmer信息
		Version: $ver

	Usage:


		-i      infile     must be given

        -o          outfile    must be given

        -h    Help document

	Usage End.

	exit;
}

###############Time
my $Time_Start;
$Time_Start = sub_format_datetime(localtime(time()));
print "\nStart Time :[$Time_Start]\n\n";
################
my $programe_dir=basename($0);
my $in=$opts{i};
my $out=$opts{o};
# my $parser   = Spreadsheet::ParseExcel->new();  #  new()生成一个新的parser  ##  这种只能读取 xls
my $parser = Spreadsheet::ParseXLSX->new;    #  new 生成一个新的parse  ，这种读取的是   xlsx的
my $workbook=$parser->parse($in);   ##  parse 存入一个新的excel文档
print "53:\t",$workbook,"\n";
if(!defined $workbook)
{
	die $parser->error(),"\n";
}

for my $worksheet( $workbook->worksheets()) { ##  得到 workbook中的所有的sheet的handler
	my ($row_min,$row_max)=$worksheet->row_range();
	my ($col_min,$col_max)=$worksheet->col_range();
	my $name = $worksheet->get_name();
	print $name,"\n"; print $row_min,"\t",$row_max,"\t";print $col_min,"\t",$col_max,"\n";
	
}

###############Time
my $Time_End;
$Time_End = sub_format_datetime(localtime(time()));
print "\nEnd Time :[$Time_End]\n\n";

###############Subs
sub sub_format_datetime {#Time calculation subroutine
    my($sec, $min, $hour, $day, $mon, $year, $wday, $yday, $isdst) = @_;
	$wday = $yday = $isdst = 0;
    sprintf("%4d-%02d-%02d %02d:%02d:%02d", $year+1900, $mon+1, $day, $hour, $min, $sec);
}

sub ABSOLUTE_DIR
{ #$pavfile=&ABSOLUTE_DIR($pavfile);
	my $cur_dir=`pwd`;chomp($cur_dir);
	my ($in)=@_;
	my $return="";
	
	if(-f $in)
	{
		my $dir=dirname($in);
		my $file=basename($in);
		chdir $dir;$dir=`pwd`;chomp $dir;
		$return="$dir/$file";
	}
	elsif(-d $in) 
	{
		chdir $in;$return=`pwd`;chomp $return;
	}
	else
	{
		warn "Warning just for file and dir in [sub ABSOLUTE_DIR]\n";
		exit;
	}
	
	chdir $cur_dir;
	return $return;
}
