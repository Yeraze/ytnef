#!/usr/bin/perl -w

use strict;
use MIME::Parser;
use MIME::Entity;
use Mail::Mailer;
use File::Path;
use Data::Dumper;

my $mail_dir = "$ENV{HOME}/.ytnef/$$";
my $reader = "/usr/bin/ytnef";
my $output_dir = "$mail_dir/output";

mkdir $mail_dir,0700;
mkdir $output_dir,0700;
my $parser = new MIME::Parser;
my $filer = new MIME::Parser::FileInto;
$filer->init($mail_dir);

$parser->filer($filer);
$parser->output_dir($mail_dir);

my $entity = $parser->parse( \*STDIN );

processParts($entity);

open(LOG, ">/dev/null");

### Delete the tnef parts of the multipart message:
my @tab;
foreach my $e ($entity->parts) {
    print LOG "working on entity:\n";
    print LOG Dumper($e);
    if ((defined $e->parts) && ($e->parts > 0)) {
        print LOG "Found subparts\n";
        foreach my $e1 ($e->parts) {
            if ( $e1->mime_type !~ /ms-tnef/i ) {
                print LOG "pushing entity:\n";
                print LOG Dumper($e1);
	        push @tab,$e1;
            }
        }
    } else {
        if ( $e->mime_type !~ /ms-tnef/i ) {
            print LOG "pushing entity:\n";
            print LOG Dumper($e);
	    push @tab,$e;
        }
    }
}
$entity->parts(\@tab);

close(LOG);

$entity->print( \*STDOUT );

$entity->purge;
$parser->filer->purge;

rmtree($mail_dir);

sub processParts {
    my $entity = shift;
    if ($entity->parts) {
        for my $part ($entity->parts) {
            processParts($part);
        }
    } else {
        if ( $entity->mime_type =~ /ms-tnef/i ) {
            if (my $bh = $entity->bodyhandle) {
                my $io = $bh->open("r") || die "Unable to open msg body";
                open(FPTR, ">$output_dir/winmail.dat");
                while (defined($_ = $io->getline)) {
                    print FPTR $_;
                }
                close(FPTR);
                $io->close;

                `$reader -f $output_dir +F $output_dir/winmail.dat`;
                unlink("$output_dir/winmail.dat");

                opendir(DIR, $output_dir) || die "Can't open directory $output_dir: $!";
                my @files = map { $output_dir."/".$_ } 
                    grep { !/^\./ }
                    readdir DIR;
                closedir DIR;

                for my $file ( @files ) {
                    my $mimetype = "application/binary";
                    my $disposition = "attachment";

                    $mimetype = "text/calendar" if ($file =~ m/\.vcf$/);
                    $mimetype = "text/x-vcard"  if ($file =~ m/\.vcard$/);
                    $mimetype = "text/html"     if ($file =~ m/\.htm[l]*$/);

                    $disposition = "inline"     if ($file =~ m/\.vcf$/);
                    $disposition = "inline"     if ($file =~ m/\.vcard$/);
					#$disposition = "inline"     if ($file =~ m/\.htm[l]*$/);

                    if ($mimetype eq "application/binary") {
                        my $qfile = quotemeta $file;
                        my $filetype = `file -bi $qfile`;
                        chomp $filetype;
                        $mimetype = $filetype if ($filetype ne "");
                    }
                    $entity->attach(
                            Path => $file,
                            Type => $mimetype,
                            Disposition => $disposition,
                            Encoding => "-SUGGEST",
                            Top => 0);
                }
            }
        }
    }
}
