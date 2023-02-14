use clap::Parser;

#[derive(Parser, Debug)]
#[clap(author="Marco Strambelli", version="0.1", about="A macro packer")]
struct Args {
    #[clap(short, long)]
    key: String,
    #[clap(long)]
    lhost: String,
    #[clap(long)]
    lport: String,
    #[clap(long)]
    rhost_path: Option<String>
}

fn main() {
   let args = Args::parse();

   println!("{:?}",args.key);
   println!("{:?}:{:?}",args.lhost,args.lport);
   println!("{:?}",args.rhost_path);
}