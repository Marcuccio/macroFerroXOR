use clap::Parser;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
   /// Key for XOR
   #[arg(short, long)]
   key: String,

   /// URL to call
   #[arg(short, long)]
   lhost: String,

   /// URL to call
   #[arg(short, long, default_value_t="8002")]
   lport: String,

   /// Number of times to greet
   #[arg(short, long, default_value_t="C:\\Users\\Public\\")]
   rhost_local_path: String,
}

fn main() {
   let args = Args::parse();



   for _ in 0..args.count {
       println!("Hello {}!", args.name)
   }
}