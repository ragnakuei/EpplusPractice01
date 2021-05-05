namespace EpplusPractice01
{
    public class Dto
    {
        public          string BB    { get; set; }
        public          string plant { get; set; }
        public          string Model { get; set; }
        public          string part  { get; set; }
        public          string pm    { get; set; }
        public          string MRP   { get; set; }
        public          string NG    { get; set; }
        public          string Total { get; set; }
        public          string type  { get; set; }
        public          string May   { get; set; }
        public          string MTDA  { get; set; }
        public          string Jun   { get; set; }

        public override string ToString()
        {
            return $@" BB: {BB} plant: {plant} Model: {Model} part: {part} pm: {pm} MRP: {MRP} NG: {NG} Total: {Total} type: {type} May: {May} MTDA: {MTDA} Jun: {Jun} ";
        }
    }
}
