#!/usr/bin/perl
use MIME::Parser;
use MIME::Entity;
use Mail::Mailer;
use Mail::Mailer::test;

my $mail_dir = "$ENV{HOME}/.ytnef";
my $reader = "/usr/bin/ytnef";
my $output_dir = "$mail_dir/output";

my $parser = new MIME::Parser;
my $filer = new MIME::Parser::FileInto;
$filer->init($mail_dir);

$parser->filer ($filer);
$parser->output_dir($mail_dir);

$entity = $parser->parse ( \*STDIN );

processParts($entity);



print STDOUT "From process.pl\n";
$entity->print( \*STDOUT );

for my $file ( @files ) {
    `rm -f $file`;
}
$entity->purge;
$parser->filer->purge;




sub processParts {
    my $entity = shift;
    if ($entity->parts) {
        for $part ($entity->parts) {
            processParts($part);
        }
    } else {
        if ( $entity->mime_type =~ /ms-tnef/i ) {
            if ($bh = $entity->bodyhandle) {
                $io = $bh->open("r");
                open(FPTR, ">$output_dir/winmail.dat");
                while (defined($_ = $io->getline)) {
                    print FPTR $_;
                }
                close(FPTR);
                $io->close;

                `$reader -f $output_dir $output_dir/winmail.dat`;
                `rm -f $output_dir/winmail.dat`;

                opendir(DIR, $output_dir) 
                    or die "Can't open directory $output_dir: $!";
                my @files = map { $output_dir."/".$_ } 
                    grep { !/^\./ }
                    readdir DIR;
                closedir DIR;

                for my $file ( @files ) {
                    #$ent = MIME::Entity->build( Path => $file,
                    #        Type => "application/binary",
                    #        Disposition => "attachment",
                    #        Encoding => "base64",
                    #        Top => 0);
                    #$entity->add_part($ent,-1);
                    $entity->attach( Path => $file,
                            Type => "application/binary",
                            Disposition => "attachment",
                            Encoding => "base64",
                            Top => 0);
                }
            }
        }
    }
}
