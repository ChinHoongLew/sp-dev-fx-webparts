import * as React from 'react';
import styles from './ReactQuickLinksFluent.module.scss';
import type { IReactQuickLinksFluentProps } from './IReactQuickLinksFluentProps';
import { Icon } from '@fluentui/react';
 
//pnp js
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/sites";
import { SPFx, spfi } from "@pnp/sp";
 

interface QuickLinkItem {
  Title: string;
  ICON: string;
  LINK: string;
  POSITION: number;
  TARGET: string;
  GROUP:string;
  COLOR:string;
  BGCOLOR:string;
}

interface IState{
  listItems: QuickLinkItem[];
  loading: boolean;
}


export default class ReactQuickLinksFluent extends React.Component<IReactQuickLinksFluentProps, IState> {

 constructor(props : IReactQuickLinksFluentProps){
  super(props);
  this.state={
    loading:true,
    listItems: [],
  };
 }

 public async componentDidMount():  Promise<undefined> {
  await this.getDataFromList();
return;
}

private async getDataFromList() :Promise<undefined>  {
console.log("Getting data from list");
try {
    
    const sp = await spfi().using(SPFx(this.props.context));

    const items = await sp.web.lists
      .getByTitle(this.props.listName)
      .items.select("Title,ICON,LINK,POSITION,TARGET,GROUP,COLOR,BGCOLOR")
      .filter(this.props.groupBy)
      .orderBy("POSITION")();

        this.setState({
      listItems: items,
      loading: false
    });}

  catch(error){console.log(error.message);}
    

    


    
    return;
  }

  public render(): React.ReactElement<IReactQuickLinksFluentProps> {

    const styleBlock = { 
      //"--tileColor": this.props.quickLinkColor,
      //"--tileColor2": this.props.quickLinkColor2,
      //"--tileFontIconColor": this.props.fontIconColor,
      "--tileMargin": this.props.margin + "px",
      "--tilePadding": this.props.padding + "px",
      "--tileMaxWidth": this.props.maxWidth + "px",
      "--tileMinHeight":this.props.minHeight+ "px",
      "--gridWidth": this.props.gridWidth+"px"
    } as React.CSSProperties;

    
    
    return (
       <div className={styles.quickLinks} style={styleBlock}>
        <div className={styles.grid}>
          {this.state.listItems.map((link, index) => {
            let boxShadowValue="";
            if (link.BGCOLOR==="transparent")
            {
               boxShadowValue ="none";
            }
            
            return(

            <div key={index} className={styles.gridItem} style={{ backgroundColor: link.BGCOLOR, boxShadow: boxShadowValue}} >
              <a data-interception="off" target={link.TARGET} href={link.LINK}>
                <Icon iconName={link.ICON} className={styles.icon} style={{ color: link.COLOR}} />
                <div style={{ color: link.COLOR}}>{link.Title}</div>
              </a>
            </div>
          

          );
        })}
        </div>
      </div>
    );
  }
}
